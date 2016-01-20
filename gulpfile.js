/**
 * Created by shuang on 2015/11/11.
 */
var gulp = require('gulp');
var zip = require('gulp-zip');
var Q = require("q");
var del = require('del');
var rev = require('gulp-rev');
var schedule = require('node-schedule');
var CronJob = require('cron').CronJob;
var fs = require("fs");
var _ = require("lodash");
var nodemailer = require('nodemailer');

var index = require('./index.js');

var config = index.AppConfig;
var logger = index.logger;
var db_report = index.db_report;
var baseFolder = config.baseFolder;

var help = function () {
    console.log('  updateAll    检查是否有新的视频，如果有，自动下载下来并打包');
    console.log('  updateVideo  手动更新视频');
    console.log('  updateSrt    手动更新字幕');
    console.log('  updateThumb  手动更新缩略图');
    console.log('  zip          将更新的视频打包');
    console.log('  cleanTmp     清空临时下载文件夹');
    console.log('  cleanSrt     删除所有的字幕文件');
    console.log('  cleanThumb   删除所有的缩略图');
}

gulp.task('default', function () {
    help();
});

gulp.task('help', function () {
    help();
});

// 手动更新视频
gulp.task('updateAll', function () {
    index.updateAll().then(function () {
        gulp.start('zip');
    });
});

gulp.task('updateVideo', function () {
    index.updateVideo();
});

gulp.task('updateSrt', function () {
    index.updateSrt();
});

gulp.task('updateThumb', function () {
    index.updateThumbnails();
})

// 打包更新的视频
gulp.task('zip', function () {
    logger.info("开始打包更新的视频");
    var today = new Date();
    var yesterday = new Date(today.getFullYear(), today.getMonth(), today.getDate()).getTime();
    var tomorrow = new Date(today.getFullYear(), today.getMonth(), today.getDate() + 1).getTime();

    var videos = [];
    for (var playlistId in db_report.object) {
        if (db_report.object.hasOwnProperty(playlistId)) {
            var arr = db_report(playlistId).filter(function (n) {
                if (n.updateAt > yesterday && n.updateAt < tomorrow) {
                    return true;
                } else {
                    return false;
                }
            })
            arr.length > 0 && (videos = videos.concat(arr));
        }
    }
    if (videos.length > 0) {
        if (videos.length > 20) {
            logger.info("更新的视频文件太多(" + videos.length + "个),无法打包。请直接拷贝整个videos目录");
            return;
        }
        var lastUpdate = today.getTime();

        var files = [];
        var tableData = [];
        var head = ['列表名称', '视频名称', '上传时间'];
        tableData.push(head);
        videos.forEach(function (video) {
            files.push(baseFolder + "/" + video.folderName + "/" + video.videoName + "*");

            var row = [];
            row.push(video.folderName);
            row.push(video.videoName);
            row.push(new Date(video.publishedAt).format("yyyyMMdd"));
            tableData.push(row);
        });
        logger.info(files);
        var sortedTableData = _.sortBy(tableData, function (n) {
            return n.publishedAt;
        })
        var tableStr = '<table style="border-collapse: collapse;">';
        for (var r = 0; r < sortedTableData.length; r++) {
            var _row = sortedTableData[r];
            if (r == 0) {
                tableStr += '<tr style="text-align: center;">';
            } else {
                tableStr += "<tr>";
            }
            for (var c = 0; c < _row.length; c++) {
                var _col = _row[c];
                tableStr += '<td style="border:1px solid #666; min-width: 100px; max-width: 300px;">' + _col + "</td>";
            }
            tableStr += "</tr>";
        }
        tableStr += "</table>";
        var contentStr = '<h3 style="margin-top: 0;">Draper TV更新了，点击链接即可下载</h3>';
        contentStr += tableStr;

        var zipName = "update_" + today.format("yyyyMMdd") + ".zip";

        gulp.src(files, {base: baseFolder})
            .pipe(zip(zipName))
            .pipe(gulp.dest(baseFolder))
            .on('end', function () {
                logger.info("打包完毕");
                db_report.object['lastUpdate'] = lastUpdate;
                db_report.saveSync();

                contentStr += '<h4>下载地址：<a href="' + config.mail.downloadAddress + zipName + '">' + zipName + '</a></h4>';
                sendMail(contentStr);
            });
    } else {
        logger.info("没有更新视频");
    }
});

var sendMail = function (contentHtml) {
    var today = new Date();
    var name = "update_" + today.format("yyyyMMdd") + ".zip";
    fs.exists(baseFolder + "/" + name, function (exists) {
        if (exists) {
            var transporter = nodemailer.createTransport({
                service: config.mail.service,
                auth: {
                    user: config.mail.user,
                    pass: config.mail.pass
                }
            });
            var mailOptions = {
                from: config.mail.user,
                to: config.mail.to,
                subject: config.mail.subject,
                html: contentHtml
            };
            transporter.sendMail(mailOptions, function (error, info) {
                if (error) {
                    return logger.error("邮件发送失败", error.stack);
                }
                logger.info('Message sent: ' + info.response);
            });
        } else {
            logger.info("打包文件不存在");
        }
    })
}

// 删除所有的字幕文件
gulp.task('cleanSrt', function () {
    del([baseFolder + '/**/*.srt', baseFolder + '/**/*.en']);
});

// 删除所有的缩略图
gulp.task('cleanThumb', function () {
    del([baseFolder + '/**/*.jpg']);
});

// 清空临时文件夹
gulp.task('cleanTmp', function () {
    del(config.tmpFolder + "/*");
});

gulp.task('package', function () {
    var now = new Date();
    var packageName = 'YouTube';
    del([packageName + "*.zip"]).then(function () {
        var resource = [
            'tmp/',
            'db/',
            'log/',
            'videos/',
            'config.json',
            'gulpfile.js',
            'main.js',
            'index.js',
            'package.json',
            'README.md'
        ];
        gulp.src(resource, {base: '../'})
            .pipe(zip('YouTube.zip'))
            .pipe(rev())
            .pipe(gulp.dest('.'));
    });
})

var startJob = function () {
    logger.info("启动定时任务", config.cron);
    new CronJob(config.cron, function () {
        logger.info("定时任务开始执行");
        gulp.start('updateAll');
    }, null, true, '');

    //schedule.scheduleJob(config.cron, function(){
    //    logger.info("定时任务开始执行");
    //    gulp.start('updateAll');
    //});
}

exports.startJob = startJob;