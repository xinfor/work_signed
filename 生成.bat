@echo off
color 0a
echo **********************************************
echo *  请保证从打卡机获取的考勤文件未更改
echo *  并改名字为 考勤报表.xls 该文件夹中
echo *                             make by XiaoY
echo **********************************************
pause
cd src
python attence.py
cls
echo **********************************************
echo *  请保证从打卡机获取的考勤文件未更改
echo *  并改名字为 考勤报表.xls 该文件夹中
echo *                             make by XiaoY
echo **********************************************
echo 完成！
echo 请检查考勤文件，进行确认是否成功
pause