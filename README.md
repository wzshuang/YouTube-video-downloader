YouTube视频下载
# 简介
程序会自动查找YouTube某一频道下的所有播放列表和视频，并下载好放到指定的文件夹内

# 安装
1. 安装nodejs 0.12x
2. 解压源程序，进入根目录
3. 修改config.json配置文件


    $ npm install -g gulp forever 
    $ npm install
    
# 启动
    $ forever stopall
    $ forever start main.js

# 备注
也可以手动运行任务,输入`gulp help`查看帮助说明  
统计视频文件总数的命令 `find . -name "*.mp4" | wc -l`


