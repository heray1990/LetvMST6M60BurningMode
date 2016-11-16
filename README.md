# Letv Max65 老化工具

此工具通过串口发送命令给 Monitor，使得 Monitor 进入老化模式。

![Maiin Form](https://github.com/heray1990/LetvMST6M60BurningMode/raw/master/Images/Main_Form.png)

## 硬件接法

串口调试小板一端连接 PC，另一端通过一根 USB 转串口的线连接到 Debug 线的 USB2.0 端口。如下图：

![Hardware Connection](https://github.com/heray1990/LetvMST6M60BurningMode/raw/master/Images/Hardware_Connection.jpg)

## 串口设置

运行此工具前，需要设置好 COM 口。进入电脑的设备管理器确认当前连接 TV 的 COM 口信息。

![USB Serial Port Setting of PC](https://github.com/heray1990/LetvMST6M60BurningMode/raw/master/Images/USB_Serial_Port_Setting.png)

然后打开工具 Setting -> COM Setting，设置串口信息，使之与 PC 的一致。

![COM Setting From](https://github.com/heray1990/LetvMST6M60BurningMode/raw/master/Images/COM_Setting_From.png)

### 运行须知

生成的 \*.exe 程序需要与 **config.xml** XML 文件和 **MSCOMM32.OCX** 文件放在同一个目录下才可以运行成功。

#### 注册 MSCOMM32.OCX

由于此工具用到了串口通讯，所以在使用之前，需要确保电脑注册了 MSCOMM32.OCX 控件。

对于 Windows 7 64 位系统，按照如下步骤进行注册：

1. 如果系统中没有 [MSCOMM32.OCX](https://github.com/heray1990/LetvMST6M60BurningMode/blob/master/MSCOMM32.OCX) 控件，需要将这个控件复制到 `C:\Windows\SysWOW64` 文件夹中。
2. 进入 `C:\Windows\SysWOW64` 文件夹，找到 `cmd.exe` 这个文件。右击，选择“以管理员身份运行”。
3. 按照下面的截图输入两条命令注册上述两个控件。

![register_ocx_in_windows7_64](https://github.com/heray1990/LetvMST6M60BurningMode/raw/master/Images/register_ocx_in_windows7_64.PNG)

> Note: 上述的 `C:\Windows\SysWOW64` 文件夹是针对 Windows 64 位系统的，对于 32 位系统，需要改成 `C:\Windows\System32`。