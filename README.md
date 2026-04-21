前提：
因每天需要给对象拍照宠物的照片，经常不能及时回复，而且宠物有时候在窝里睡觉，懒得起身去拍照，所以基于自动发送，使用ai开发了监听关键词进行回复的软件
1、建议使用虚拟机运行


2、打开聊天软件前要打开讲述人，登录后可以关闭（注意是打开软件前打开。因为这样才能激活4.0以上的控件树）


3、如果不稳定或者识别失效可以换成本地OCR方式




4、建议把软件对话框独立出来并且置顶


5、如果用rdp远程桌面，退出远程不能直接关闭窗口，会导致没实体屏幕，控件树不能生效，需要用到命令行执行：@echo off
for /f "skip=1 tokens=3" %%s in ('query user %USERNAME%') do ( %windir%\System32\tscon.exe %%s /dest:console )保存为bat，需要退出远程桌面时执行


顺便做了个软件，适配鸿蒙6.0系统，在HarmonyOS-NAS-Photo-Sync1仓库下，


现在有空可以手机拍摄不同宠物照片素材，并上传到nas，设定软件监控，当到对方发送宠物名字，就会随机发送宠物照片，并且设定随机延迟，模拟真人发送


# EasyChat Momo

当前仓库主入口为 `wechat_gui_momo.py`，用于微信关键词触发的图片/文本自动回复。

运行方式：

```bash
py wechat_gui_momo.py
```

打包方式：

```bash
py pack.py
```

生成便携包：

```bash
py pack.py --portable
```

![preview](https://github.com/609293964/easyChat--/blob/main/pictures/111.png?raw=true)
