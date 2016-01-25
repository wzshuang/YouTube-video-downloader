var $ = require('cheerio');
var _ = require('lodash');
var request = require('superagent');
var request_node = require('request');
var ProgressBar = require('progress');
var log4js = require('log4js');
var fs = require("fs");
var Q = require("q");
var low = require('lowdb');
var xlsx = require('node-xlsx');
var AppConfig = require('./config.json');

var db_playlists = low('db/playlists.json', {async: false});
var db_playlistitems = low('db/playlistitems.json', {async: false});
var db_downloads = low('db/downloads.json', {async: false});
var db_report = low('db/report.json', {async: false});
var db_uploadvideos = low('db/uploadvideos.json', {async: false});

log4js.configure('config.json', {});
var logger = log4js.getLogger('main');

var channelId = AppConfig.channelId;
var apiKey = AppConfig.apiKey;
var baseFolder = AppConfig.baseFolder;
var tmpFolder = AppConfig.tmpFolder;
var uploadFolder = "upload";
var zhSrtFolder = "zh-srt";
var uploadPlaylistId = "UU" + channelId.substring(2);

// 日期格式化 yyyy-MM-dd HH:mm:ss:S
Date.prototype.format = function (fmt) {
    var o = {
        "M+": this.getMonth() + 1, //月份
        "d+": this.getDate(), //日
        "h+": this.getHours() % 12 == 0 ? 12 : this.getHours() % 12, //小时
        "H+": this.getHours(), //小时
        "m+": this.getMinutes(), //分
        "s+": this.getSeconds(), //秒
        "q+": Math.floor((this.getMonth() + 3) / 3), //季度
        "S": this.getMilliseconds() //毫秒
    };
    var week = {
        "0": "/u65e5",
        "1": "/u4e00",
        "2": "/u4e8c",
        "3": "/u4e09",
        "4": "/u56db",
        "5": "/u4e94",
        "6": "/u516d"
    };
    if (/(y+)/.test(fmt)) {
        fmt = fmt.replace(RegExp.$1, (this.getFullYear() + "").substr(4 - RegExp.$1.length));
    }
    if (/(E+)/.test(fmt)) {
        fmt = fmt.replace(RegExp.$1, ((RegExp.$1.length > 1) ? (RegExp.$1.length > 2 ? "/u661f/u671f" : "/u5468") : "") + week[this.getDay() + ""]);
    }
    for (var k in o) {
        if (new RegExp("(" + k + ")").test(fmt)) {
            fmt = fmt.replace(RegExp.$1, (RegExp.$1.length == 1) ? (o[k]) : (("00" + o[k]).substr(("" + o[k]).length)));
        }
    }
    return fmt;
}
String.prototype.format = function () {
    if (arguments.length == 0) return this;
    for (var s = this, i = 0; i < arguments.length; i++) {
        s = s.replace(new RegExp("\\{" + i + "\\}", "g"), arguments[i]);
    }
    return s;
};

process.on('uncaughtException', function (err) {
    //打印出错误
    logger.error("严重：出现全局异常！");
    //打印出错误的调用栈方便调试
    logger.error(err.stack);
});

var setRequest = function (_request) {
    _request
        .set('User-Agent', 'Mozilla/5.0 (Windows NT 6.3; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/46.0.2490.80 Safari/537.36')
        .set('Upgrade-Insecure-Requests', 1)
        .set('Referer', 'https://content.googleapis.com/static/proxy.html?jsh=m%3B%2F_%2Fscs%2Fapps-static%2F_%2Fjs%2Fk%3Doz.gapi.zh_CN.tkiNpjG4rhI.O%2Fm%3D__features__%2Fam%3DAQ%2Frt%3Dj%2Fd%3D1%2Ft%3Dzcms%2Frs%3DAGLTcCNhbBDeEP99y0jGa8pLoqSI515feQ')
        .set('X-Referer', 'https://developers.google.com')
        .set('X-Origin', 'https://developers.google.com')
        .set('X-ClientDetails', 'appVersion=5.0%20(Windows%20NT%206.3%3B%20WOW64)%20AppleWebKit%2F537.36%20(KHTML%2C%20like%20Gecko)%20Chrome%2F46.0.2490.80%20Safari%2F537.36&platform=Win32&userAgent=Mozilla%2F5.0%20(Windows%20NT%206.3%3B%20WOW64)%20AppleWebKit%2F537.36%20(KHTML%2C%20like%20Gecko)%20Chrome%2F46.0.2490.80%20Safari%2F537.36')
        .set('X-Goog-Encode-Response-If-Executable', 'base64')
        .set('X-JavaScript-User-Agent', 'google-api-javascript-client/1.1.0-beta');
    return _request;
};

//用户上传的所有视频都会出现在“视频”标签下，“播放列表”可以包含“视频”里的视频，同一个视频可以属于多个播放列表。

// 获取playlists
var getPlaylists = function (nextPageToken) {
    var deferred = Q.defer();
    var listArr = [];
    var inner = function (nextPageToken) {
        var _req = request.get('https://content.googleapis.com/youtube/v3/playlists');
        setRequest(_req)
            .query({
                'part': 'snippet',
                'maxResults': 50,
                'channelId': channelId,
                'key': apiKey,
                'pageToken': nextPageToken
            })
            .end(function (err, res) {
                if (res !== undefined && res.ok) {
                    logger.info("获取playlists...");
                    listArr = listArr.concat(res.body.items);
                    var nextPageToken = res.body.nextPageToken;
                    if (nextPageToken !== undefined) {
                        logger.info("获取下一页的playlists...");
                        inner(nextPageToken);
                    } else {
                        logger.info("playlists获取完毕...");
                        db_playlists.object[channelId] = listArr;
                        db_playlists.saveSync();
                        deferred.resolve(channelId);
                    }
                } else {
                    logger.error("网络无法连接或服务出错");
                    deferred.reject(res);
                }
            });
    }
    inner();
    return deferred.promise;
}

// 获取playlistItems
var getPlaylistItems = function () {
    var _deferred = Q.defer();
    var inner = function (playlistId, nextPageToken) {
        var deferred = Q.defer();
        var _req = request.get('https://content.googleapis.com/youtube/v3/playlistItems');
        setRequest(_req)
            .query({
                'part': 'snippet',
                'maxResults': 50,
                'playlistId': playlistId,
                'key': apiKey,
                'pageToken': nextPageToken
            })
            .end(function (err, res) {
                if (res !== undefined && res.ok) {
                    db_playlistitems.object[playlistId] = _.filter(res.body.items, function (n) {
                        // 私有的视频没有权限看，跳过
                        return (n.snippet.title != 'Private video' && n.snippet.resourceId.videoId !== undefined);
                    });
                    db_playlistitems.saveSync();
                    var nextPageToken = res.body.nextPageToken;
                    if (nextPageToken !== undefined) {
                        //inner(playlistId, nextPageToken);
                    } else {

                    }
                    deferred.resolve();
                } else {
                    logger.error("网络错误");
                    deferred.resolve();
                }
            });
        return deferred.promise;
    }
    logger.info("读取已经加载的playlists...");
    var qArr = [];
    db_playlists(channelId).forEach(function (list) {
        qArr.push(inner(list.id));
    });
    Q.allSettled(qArr).then(function () {
        logger.info("playlistitems获取完毕...");
        _deferred.resolve();
    });
    return _deferred.promise;
}

// 获取上传的视频
var getUploadVideos = function () {
    logger.info("开始获取所有的上传视频...");
    var _deferred = Q.defer();
    var uploadVideos = [];
    var inner = function (nextPageToken) {
        var _req = request.get('https://content.googleapis.com/youtube/v3/playlistItems');
        setRequest(_req)
            .query({
                'part': 'snippet',
                'maxResults': 50,
                'playlistId': uploadPlaylistId,
                'key': apiKey,
                'pageToken': nextPageToken
            })
            .end(function (err, res) {
                if (res !== undefined && res.ok) {
                    //logger.info("将playlistitems写入文件..." + playlistId);
                    var v = _.filter(res.body.items, function (n) {
                        return (n.snippet.title != 'Private video' && n.snippet.resourceId.videoId !== undefined);
                    });
                    if (v.length > 0) {
                        uploadVideos = uploadVideos.concat(v);
                    }
                    var nextPageToken = res.body.nextPageToken;
                    if (nextPageToken !== undefined) {
                        //logger.info("获取下一页的playlistitems...");
                        inner(nextPageToken);
                    } else {
                        logger.info("获取所有的上传视频完毕");
                        uploadVideos.forEach(function (video) {
                            var videoId = video.snippet.resourceId.videoId;
                            for (var _playlistId in db_playlistitems.object) {
                                if (db_playlistitems.object.hasOwnProperty(_playlistId)) {
                                    db_playlistitems.object[_playlistId].forEach(function (item) {
                                        if (item.snippet.resourceId.videoId == videoId) {
                                            video.inPlaylist = true;
                                        }
                                    })
                                }
                            }
                        });

                        var _arr = _.remove(uploadVideos, function (video) {
                            return video.inPlaylist != true;
                        });

                        db_uploadvideos.object[channelId] = _arr;
                        db_uploadvideos.saveSync();

                        _deferred.resolve();
                    }
                } else {
                    logger.error("网络错误");
                    deferred.resolve();
                }
            });
    }
    inner();

    return _deferred.promise;
}

// windows下的文件名不能包含某些特殊字符。修改此方法将导致旧数据出现异常
var replaceFileName = function (name) {
    name = name.replace(/:/g, "：");
    name = name.replace(/\|/g, "-");
    name = name.replace(/\?/g, "？");
    name = name.replace(/</g, "《");
    name = name.replace(/>/g, "》");
    name = name.replace(/\"/g, "”");

    name = name.replace(/\&"/g, "＆");
    name = name.replace(/\/"/g, "＆");
    return name;
}

var _mkdir = function (path, index) {
    var defer = Q.defer();
    logger.info("开始创建文件夹", path);
    fs.exists(path, function (exists) {
        if (exists) {
            logger.info("文件夹已存在，略过", path);
            defer.resolve();
        } else {
            // 延时是为了制造创建时间差，这样就可以根据创建时间排序了
            setTimeout(function () {
                fs.mkdir(path, function (err) {
                    if (err) {
                        logger.error("创建文件夹失败", err.stack);
                    } else {
                        logger.info("创建文件夹成功", path);
                    }
                    defer.resolve();
                })
            }, 1000 * (index || 1));
        }
    });
    return defer.promise;
}
// 根据播放列表的名称创建文件夹
var createFolder = function () {
    var defer = Q.defer();
    var lists = db_playlists(channelId).filter();
    _mkdir(baseFolder).then(function () {
        var uploadFolderName = baseFolder + "/" + uploadFolder;
        return _mkdir(uploadFolderName).then(function () {
            return _mkdir(uploadFolderName + "/" + zhSrtFolder);
        })
    }).then(function () {
        var qArr = [];
        var index = 0;
        lists.forEach(function (list) {
            var title = list.snippet.title;
            var publishedAt = new Date(list.snippet.publishedAt).format("-yyyyMMdd");
            var folderName = baseFolder + '/' + replaceFileName(title) + publishedAt;
            var p = _mkdir(folderName, ++index).then(function () {
                return _mkdir(folderName + "/" + zhSrtFolder, 0);
            })
            qArr.push(p);
        })
        Q.allSettled(qArr).then(function () {
            logger.info("文件夹创建完毕...");
            defer.resolve();
        });
    });
    return defer.promise;
}

// 工具方法，下载文件，保存到指定的目录
var _downloadFile = function (folderName, fileName, url) {
    logger.info("准备下载文件", fileName);
    var _defer = Q.defer();
    var path = folderName + "/" + fileName;
    if (!url) {
        logger.error("文件下载地址错误，略过", fileName, url);
        _defer.reject();
    } else {
        var receivedSize, totalSize, startTime, endTime;
        request_node.get(url)
            .on('response', function (response) {
                logger.info("开始下载文件", fileName);
                startTime = new Date().getTime();
                totalSize = Number(response.headers['content-length']);
                receivedSize = 0;
                var kb = Number(totalSize / 1024).toFixed(1);
                var mb = Number(kb / 1024).toFixed(1);
                var size = kb + "KB";
                kb > 10000 && (size = mb + "MB");
                !isNaN(totalSize) && logger.info("文件总大小", size);
                bar = new ProgressBar('  下载进度 [:bar] :percent', {
                    complete: '=',
                    incomplete: ' ',
                    width: 40,
                    total: totalSize || 0
                });
            })
            .on('data', function (data) {
                receivedSize += data.length;
                !isNaN(totalSize) && bar.tick(data.length);
                !process.stderr.isTTY && process.stdout.write(fileName + '  下载进度 ' + Number(receivedSize * 100 / totalSize).toFixed(0) + '% \r');
            })
            .on('close', function () {
                logger.info('close event fired');
                //_defer.resolve(true);
            })
            .on('end', function (err, res) {
                if (receivedSize && receivedSize > 500) {
                    endTime = new Date().getTime();
                    var seconds = (endTime - startTime) / 1000;
                    logger.info("文件下载成功, 耗时" + Number(seconds / 60).toFixed(0) + "分" + Number(seconds % 60).toFixed(0) + "秒 " + fileName);
                    try {
                        var source = fs.createReadStream(tmpFolder + '/' + fileName);
                        var dest = fs.createWriteStream(folderName + "/" + fileName);
                        source.pipe(dest);
                        source.on('end', function () {
                            fs.unlink(tmpFolder + '/' + fileName);
                            // 延时后创建时间不一样，可以用来根据创建时间排序
                            setTimeout(function () {
                                _defer.resolve();
                            }, 1000);
                        });
                        source.on('error', function (err) {
                            logger.error(err.stack);
                            _defer.reject();
                        });
                    } catch (e) {
                        logger.error(e.stack);
                        _defer.reject();
                    }
                } else {
                    logger.error("文件下载失败 ", fileName, receivedSize);
                    _defer.reject();
                }
            })
            .pipe(fs.createWriteStream(tmpFolder + '/' + fileName))
            .on('error', function (err) {
                logger.error(err.stack);
                _defer.reject();
            });
    }
    return _defer.promise;
}

// 采访人在哪个播放列表里。如果视频不在播放列表里，但是采访人的名字在某个播放列表出现过，
// 那就把这个视频归到那个播放列表
var _getUserPlaylist = function () {
    var userPlaylist = {};
    for (var pId in db_playlistitems.object) {
        if (db_playlistitems.object.hasOwnProperty(pId)) {
            db_playlistitems.object[pId].forEach(function (item) {
                var itemName = item.snippet.title;
                var userName = itemName.substring(itemName.lastIndexOf("|") + 1).trim();
                userPlaylist[userName] = item.snippet.playlistId;
            })
        }
    }
    return userPlaylist;
}

// 从数据库获取播放列表的名称、视频的名称、视频的id, 返回的是数组
var _getVideosInfoForDownload = function () {
    var arr = [];

    var userPlaylist = _getUserPlaylist();
    db_uploadvideos.object[channelId].forEach(function (video) {
        var userName = video.snippet.title.substring(video.snippet.title.lastIndexOf("|") + 1).trim();
        var pId = userPlaylist[userName];
        var obj = {
            videoId: video.snippet.resourceId.videoId,
            videoName: replaceFileName(video.snippet.title),
            publishedAt: video.snippet.publishedAt,
            position: video.snippet.position
        };
        if (pId != undefined) {
            var playlist = db_playlists(channelId).find({id: pId});
            var playlistName = replaceFileName(playlist.snippet.title);
            obj.playlistId = pId;
            obj.folderName = playlistName;
            obj.folderName = playlistName + new Date(playlist.snippet.publishedAt).format("-yyyyMMdd");
        } else {
            obj.playlistId = uploadPlaylistId;
            obj.playlistName = uploadFolder;
            obj.folderName = uploadFolder;
        }
        arr.push(obj);
    });

    db_playlists(channelId).filter().forEach(function (list) {
        var listitems = db_playlistitems.object[list.id];
        listitems.forEach(function (item) {
            var playlistName = replaceFileName(list.snippet.title);
            arr.push({
                playlistId: list.id,
                playlistName: playlistName,
                folderName: playlistName + new Date(list.snippet.publishedAt).format("-yyyyMMdd"),
                videoName: replaceFileName(item.snippet.title),
                videoId: item.snippet.resourceId.videoId,
                position: item.snippet.position,
                publishedAt: item.snippet.publishedAt
            })
        })
    });

    return _.sortBy(arr, 'publishedAt');
}

// 根据视频id获取真实的下载地址
var requestVideoDownloadUrl = function (videoId) {
    var defer = Q.defer();
    logger.info("开始获取视频的下载地址", videoId);
    request
        .get('http://www.youtube.com/get_video_info')
        .set('User-Agent', 'Mozilla/5.0 (Windows NT 6.3; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/46.0.2490.80 Safari/537.36')
        .query({video_id: videoId})
        .end(function (err, res) {
            if (err) {
                logger.error(err.stack);
                return;
            }
            var data = res.body;
            var map = data.url_encoded_fmt_stream_map.split("&");
            var url;
            for (var i = 0; i < map.length; i++) {
                var data = map[i].split("=");
                if (data[0] == "url") {
                    url = decodeURIComponent(data[1]);
                    logger.info("获取视频的下载地址成功", videoId);
                    defer.resolve(url);
                    break;
                }
            }
        })
    return defer.promise;
}

// 根据视频id获取字幕的下载地址
var requestSrtDownloadUrl = function (videoId) {
    var defer = Q.defer();
    logger.info("开始获取字幕", videoId);
    try {
        request
            .get('http://downsub.com')
            .set('User-Agent', 'Mozilla/5.0 (Windows NT 6.3; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/46.0.2490.80 Safari/537.36')
            .set('Referer', 'http://downsub.com/')
            .set('Host', 'downsub.com')
            .query({'url': 'http://youtube.com/watch?v=' + videoId})
            .end(function (err, res) {
                var obj;
                if (!err || !res.ok) {
                    var result = JSON.parse(res.text);
                    if (result.state == 2) {
                        logger.info("获取字幕成功", videoId);
                        obj = {lang: {}};
                        obj.videoId = videoId;
                        obj.title = replaceFileName(result.title);
                        result.autotrans && result.autotrans.forEach(function (lang) {
                            if (lang.n == "en" || lang.n == "zh-Hans") {
                                obj.lang[lang.n] = lang.url;
                            }
                        })
                    } else {
                        logger.error("没有字幕", videoId);
                    }
                } else {
                    logger.error('获取字幕失败' + err.stack);
                }
                defer.resolve(obj);
            });
    } catch (e) {
        logger.error(e.stack);
        defer.resolve();
    }
    return defer.promise;
}

// 下载视频
var downloadVideos = function () {
    logger.info("开始下载视频...");
    var defer = Q.defer();
    var index = 0;
    var data = _getVideosInfoForDownload();
    var inner = function (index) {
        logger.info("视频下载进度：" + (index + 1), data.length);
        var next = function () {
            if (index != data.length - 1) {
                inner(++index);
            } else {
                logger.info("视频下载完毕...");
                defer.resolve();
            }
        }
        var param = data[index];
        var suffix = ".mp4";

        fs.exists(baseFolder + "/" + param.folderName + "/" + param.videoName + suffix, function (exists) {
            if (exists) {
                logger.info("视频已存在，略过", param.videoName);
                next();
            } else {
                var retryCount = 3;
                var retryDownload = function () {
                    requestVideoDownloadUrl(param.videoId).then(function (url) {

                        return _downloadFile(baseFolder + "/" + param.folderName, param.videoName + suffix, url);
                    }).then(function () {
                        var r = db_report(param.playlistId).find({videoId: param.videoId});
                        if (!r) {
                            db_report(param.playlistId).push({
                                videoId: param.videoId,
                                folderName: param.folderName,
                                videoName: param.videoName,
                                publishedAt: param.publishedAt
                            });
                            r = db_report(param.playlistId).find({videoId: param.videoId});
                        }
                        r['updateAt'] = new Date().getTime();
                        db_report.saveSync();
                        next();
                    }).catch(function () {
                        if (retryCount-- > 0) {
                            logger.info("视频下载失败, 重试下载", param.videoName);
                            retryDownload();
                        } else {
                            logger.error("视频下载失败, 跳过", param.videoName);
                            next();
                        }
                    })
                }
                retryDownload();
            }
        })
    }
    inner(index);
    return defer.promise;
}

// 下载字幕
var downloadSrts = function () {
    var defer = Q.defer();
    logger.info("开始下载字幕...");
    var index = 0;
    var data = _getVideosInfoForDownload();
    var inner = function (index) {
        var next = function () {
            logger.info("字幕下载进度：" + (index + 1), data.length);
            if (index != data.length - 1) {
                inner(++index);
            } else {
                logger.info("字幕下载完毕");
                defer.resolve();
            }
        }
        var param = data[index];
        var en_suffix = ".srt";
        var zh_suffix = ".srt";
        var tmpPath = baseFolder + "/" + param.folderName + "/" + param.videoName;

        fs.exists(tmpPath + en_suffix, function (en_exists) {
            fs.exists(tmpPath + zh_suffix, function (zh_exists) {
                if (en_exists && zh_exists) {
                    //logger.info("字幕已存在，略过");
                    next();
                } else {
                    var retryCount = 3;
                    var retryDownload = function () {
                        requestSrtDownloadUrl(param.videoId).then(function (srtInfo) {
                            if ((typeof srtInfo) === "undefined") {
                                next();
                            } else {
                                var url_en = 'http://downsub.com/index.php?title=' + param.videoName + '&url=' + encodeURIComponent(srtInfo['lang']['en']);
                                var url_zh = 'http://downsub.com/index.php?title=' + param.videoName + '&url=' + encodeURIComponent(srtInfo['lang']['zh-Hans']);
                                _downloadFile(baseFolder + "/" + param.folderName, param.videoName + en_suffix, url_en).then(function () {
                                    return _downloadFile(baseFolder + "/" + param.folderName + "/" + zhSrtFolder, param.videoName + zh_suffix, url_zh)
                                }).then(function () {
                                    next();
                                }).catch(function () {
                                    if (retryCount-- > 0) {
                                        logger.info("字幕下载失败, 重试下载", param.videoName);
                                        retryDownload();
                                    } else {
                                        logger.error("字幕下载失败, 跳过", param.videoName);
                                        next();
                                    }
                                });
                            }
                        });
                    };
                    retryDownload();
                }
            })
        });
    };
    inner(index);
    return defer.promise;
};

var _getThumbnailsInfoForDownload = function () {

};
// 下载缩略图
var downloadThumbnails = function () {
    var defer = Q.defer();
    var suffix = ".jpg";
    var videos = _getVideosInfoForDownload();
    var index = 0;
    var inner = function (i) {
        logger.info("开始下载缩略图", i, videos.length);
        var next = function () {
            if (index != videos.length - 1) {
                inner(++index);
            } else {
                logger.info("缩略图下载完毕...");
                defer.resolve();
            }
        };
        var video = videos[i];
        var videoId = video.videoId;
        var url = "https://i.ytimg.com/vi/" + videoId + "/hqdefault.jpg";
        var retryCount = 3;
        var retryDownload = function () {
            fs.exists(baseFolder + "/" + video.folderName + "/" + video.videoName + suffix, function (exists) {
                if (exists) {
                    logger.info("视频缩略图已存在", video.videoName);
                    next();
                } else {
                    _downloadFile(baseFolder + "/" + video.folderName, video.videoName + suffix, url).then(function () {
                        if (video.position == 0) {
                            try {
                                var source = fs.createReadStream(baseFolder + "/" + video.folderName + '/' + video.videoName + suffix);
                                var dest = fs.createWriteStream(baseFolder + "/" + video.folderName + '/' + video.playlistName + suffix);
                                source.pipe(dest);
                                source.on('end', function () {
                                    next();
                                });
                            } catch (e) {
                                logger.error("下载播放列表缩略图失败2", video.playlistName, e.stack);
                                next();
                            }
                        } else {
                            next();
                        }
                    }).catch(function () {
                        if (retryCount-- > 0) {
                            logger.info("播放列表缩略图下载失败, 重试下载", video.playlistName);
                            retryDownload();
                        } else {
                            logger.error("播放列表缩略图下载失败, 跳过", video.playlistName);
                            next();
                        }
                    })
                }
            })
        };
        retryDownload();
        if (video.position === 0) {
        }
    };
    inner(index);
    return defer.promise;
};

var generateExcel = function () {
    logger.info("开始生成excel表格...");
    var defer = Q.defer();
    var videos = {};
    for (var pId in db_playlistitems.object) {
        if (db_playlistitems.object.hasOwnProperty(pId)) {
            var items = db_playlistitems.object[pId];
            items && items.forEach(function (item) {
                videos[item.snippet.resourceId.videoId] = item;
            })
        }
    }
    db_uploadvideos.object[channelId].forEach(function (video) {
        videos[video.snippet.resourceId.videoId] = video;
    });
    var videosArr = [];
    for (var id in videos) {
        if (videos.hasOwnProperty(id)) {
            videosArr.push(videos[id]);
        }
    }
    var data = [];
    var head = ['列表', '上传时间', '列表简介', '视频标题', '视频简介', '采访人', 'URL'];
    data.push(head);
    _.sortBy(videosArr, 'snippet.publishedAt').forEach(function (n) {
        var playlistId = n.snippet.playlistId;
        var playlist = db_playlists(channelId).find({id: playlistId});
        var playlistName = "";
        var playlistDescription = "";
        if (playlist !== undefined) {
            playlistName = playlist.snippet.title;
            playlistDescription = playlist.snippet.description;
        }
        var videoName = n.snippet.title;
        var interViewer = "";
        var index = videoName.lastIndexOf("|");
        if (index > -1) {
            interViewer = videoName.substring(index + 1).trim();
        }
        var row = [];
        // 播放列表名称
        row.push(playlistName);
        // 视频上传时间
        row.push(new Date(n.snippet.publishedAt).format("yyyyMMdd") * 1);
        // 播放列表简介
        row.push(playlistDescription);
        // 视频名称
        row.push(videoName);
        // 视频简介
        row.push(n.snippet.description);
        // 采访者名称
        row.push(interViewer);
        // 视频地址
        row.push("http://www.youtube.com/watch?v=" + n.snippet.resourceId.videoId);
        data.push(row);
    });
    var buffer = xlsx.build([{name: AppConfig.excelName, data: data}]);
    fs.open(AppConfig.excelName + '.xlsx', 'w', function (err, fd) {
        if (err) {
            logger.error("打开excel表格失败", err.stack);
            defer.resolve(false);
        }
        fs.write(fd, buffer, 0, buffer.length, null, function (err) {
            if (err) {
                logger.error("生成excel表格失败", err.stack);
                defer.resolve(false);
            }
            fs.close(fd, function () {
                logger.info("生成excel表格成功");
                defer.resolve(true);
            })
        });
    });
    return defer.promise;
};

var updateAll = function () {
    var defer = Q.defer();
    getPlaylists().then(function () {
        return getPlaylistItems();
    }).then(function () {
        return getUploadVideos();
    }).then(function () {
        return generateExcel();
    }).then(function () {
        return createFolder();
    }).then(function () {
        return downloadVideos();
    }).then(function () {
        return downloadSrts()
    }).then(function () {
        return downloadThumbnails();
    }).then(function () {
        logger.info("全部更新完毕");
        defer.resolve();
    });
    return defer.promise;
}

var updateVideo = function () {
    var defer = Q.defer();
    getPlaylists().then(function () {
        return getPlaylistItems();
    }).then(function () {
        return getUploadVideos();
    }).then(function () {
        return generateExcel();
    }).then(function () {
        return createFolder();
    }).then(function () {
        return downloadVideos();
    });
    return defer.promise;
}

var updateSrt = function () {
    var defer = Q.defer();
    getPlaylists().then(function () {
        return getPlaylistItems();
    }).then(function () {
        return getUploadVideos();
    }).then(function () {
        return generateExcel();
    }).then(function () {
        return createFolder();
    }).then(function () {
        return downloadSrts();
    });
    return defer.promise;
}

var updateThumbnails = function () {
    var defer = Q.defer();
    getPlaylists().then(function () {
        return getPlaylistItems();
    }).then(function () {
        return getUploadVideos();
    }).then(function () {
        return generateExcel();
    }).then(function () {
        return createFolder();
    }).then(function () {
        return downloadThumbnails();
    });
    return defer.promise;
};

var renameFolder = function () {
    var defer = Q.defer();
    var playlists = db_playlists(channelId).filter();
    var index = 0;
    var inner = function (i) {
        console.log(playlists[i]);
        var publishedAt = new Date(playlists[i].snippet.publishedAt).format("yyyyMMdd-");
        var title = replaceFileName(playlists[i].snippet.title);
        fs.rename(baseFolder + "/" + title, baseFolder + "/" + publishedAt + title, function (err) {
            console.log(".");
        });
        index++;
        if (index < playlists.length) {
            inner(index);
        } else {
            defer.resolve();
        }
    }
    inner(index);
    return defer.promise;
}

exports.logger = logger;
exports.db_report = db_report;
exports.AppConfig = AppConfig;
exports.updateAll = updateAll;
exports.updateVideo = updateVideo;
exports.updateSrt = updateSrt;
exports._getVideosInfoForDownload = _getVideosInfoForDownload;
exports.updateThumbnails = updateThumbnails;


