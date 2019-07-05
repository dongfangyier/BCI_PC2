using System;
using System.Collections.Generic;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace BCI_PC
{
    class BCIEntity
    {
        public SerialPort port = new SerialPort();
        private List<byte> rawBuffer;
        private int CmdNum = -1;  // 消息序号

        public BCIEntity(string name = "COM7", int rate = 921600)
        {
            if (port.IsOpen)
            {
                port.Close();
            }
            port.PortName = name;
            port.BaudRate = rate;
            port.DataBits = 8;
            port.StopBits = StopBits.One;
            port.Parity = 0;
            CmdNum = -1;
            port.Open();
            port.DataReceived += Port_DataReceived;
        }

        #region 接收数据
        private void Port_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            var sp = sender as SerialPort;
            var length = sp.BytesToRead;
            var buf = new byte[length];
            try
            {
                sp.Read(buf, 0, length);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error reading serial data! {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            rawBuffer.AddRange(buf);

            DataParser();

        }

        private void DataParser()
        {
            string[] bodybuf = new string[33];

            while (rawBuffer.Count >= 5)
            {
            }

        }

        #endregion


        #region 处理命令
        // 添加校验和
        private string CalCheckSum(string AnteriorSegment)
        {
            int len = AnteriorSegment.Length - BCICommand.CMD_FRAME_HEAD.Length;
            len = (len / 2 * 8) % 256;
            string s = len.ToString("X2"); // 转为16进制
            return AnteriorSegment+ "\\x" + s;
        }

        // 获得消息序号
        private string GetCmdNum()
        {
            CmdNum++;
            if (CmdNum == 255)
            {
                CmdNum = 0;
            }
            return "\\x" + CmdNum.ToString("X2");
        }

        // 给命令添加消息序号
        private string AddCmdNum(string AnteriorSegment)
        {
            int pos = (2 + 2 + 1) * 4;
            AnteriorSegment.Insert(4, GetCmdNum());
            return AnteriorSegment;
        }

        // 添加序号和校验和
        // 一般情况下这个板块只需要调用这一个函数
        private string HandleCMD(string AnteriorSegment)
        {
            AnteriorSegment = AddCmdNum(AnteriorSegment);
            AnteriorSegment = CalCheckSum(AnteriorSegment);
            return AnteriorSegment;
        }

        #endregion

        #region 命令
        // 通道开控制
        private void OpenChannel()
        {
            port.Write(HandleCMD(BCICommand.OPEN_CHANNEL));
        }

        // 通道关
        private void CloseChannel()
        {
            port.Write(HandleCMD(BCICommand.CLOSE_CHANNEL));
        }

        // 连接到测试通道
        // channelType说明：
        //0：连接到内部 GND(VDD-VSS)
        //1：连接到测试信号 1倍幅度的慢脉冲
        //2：连接到测试信号 1倍幅度的快脉冲
        //3：连接到直流信号
        //4：连接到测试信号 2倍幅度的慢脉冲
        //5：连接到测试信号 2倍幅度的快脉冲
        private void TestChannel(int channel, int channelType)
        {
            if (channel < 0 || channel > 21)
            {
                MessageBox.Show("通道号有错误！");
                return;
            }
            if (channelType < 0 || channelType > 5)
            {
                MessageBox.Show("通道类型有错误！");
                return;
            }
            string cmd = BCICommand.TEST_CHANNEL + "\\x" + channel.ToString("X2") + "\\x" + channelType.ToString("X2");
            port.Write(HandleCMD(cmd));

        }

        private string HandleLen(string res, int len)
        {
            if(res.Length == len)
            {
                return res;
            }
            else
            {
                int dvalue = len - res.Length;
                for(int i = 0; i < dvalue; i++)
                {
                    res = '0' + res;
                }
                return res;
            }
        }

        // 通道设置
        private void SetChannel(int CHANNEL, int BATT = 0, int SRB1 = 1, int SRB2 = 1, int BIAS = 1, int TYPE = 0,int GAN = 6)
        {
            if (CHANNEL < 0 || CHANNEL > 21)
            {
                MessageBox.Show("通道号有错误！");
                return;
            }

            string cmd = BCICommand.SET_CHANNEL;
            string channel = Convert.ToString(CHANNEL, 2);
            channel = HandleLen(channel, 5);
            string srb1 = Convert.ToString(SRB1, 2);
            string srb2 = Convert.ToString(SRB2, 2);
            string byte1 = channel + srb1 + srb2;
            byte1= "//x" + Convert.ToInt32(byte1, 2).ToString("X2");

            string bias = Convert.ToString(BIAS, 2);
            string type = Convert.ToString(TYPE, 2);
            type = HandleLen(type, 3);
            string gan = Convert.ToString(GAN, 2);
            gan = HandleLen(gan, 3);
            string batt = Convert.ToString(BATT, 2);
            string byte2 = bias + type + gan+batt;
            byte2 = "//x" + Convert.ToInt32(byte2, 2).ToString("X2");

            cmd = cmd + byte1 + byte2;

            port.Write(HandleCMD(cmd));

        }


        // 阻抗测试
        private void TestImpedance(int CHANNEL, int REVERSAL = 0, int N = 0, int P = 0)
        {
            if (CHANNEL < 0 || CHANNEL > 21)
            {
                MessageBox.Show("通道号有错误！");
                return;
            }
            string cmd = BCICommand.TEST_IMPEDANCE;
            string channel = Convert.ToString(CHANNEL, 2);
            channel = HandleLen(channel, 5);
            string reversal = Convert.ToString(REVERSAL, 2);
            string n = Convert.ToString(N, 2);
            string p = Convert.ToString(P, 2);
            string byte1 = channel + reversal + n + p;
            byte1 = "//x" + Convert.ToInt32(byte1, 2).ToString("X2");

            cmd = cmd + byte1;

            port.Write(HandleCMD(cmd));
        }


        // 采集控制 start
        private void AcqControlStart(int MODE=0,int RATE=6,int START=1)
        {
            string cmd = BCICommand.ACQUISTION_CONTROL;

            string mode = Convert.ToString(MODE, 2);
            mode = HandleLen(mode, 3);
            string rate = Convert.ToString(RATE, 2);
            rate = HandleLen(rate, 3);
            string start = Convert.ToString(START, 2);
            string byte1 = '0' + mode + rate + start;
            byte1 = "//x" + Convert.ToInt32(byte1, 2).ToString("X2");
            cmd = cmd + byte1;

            port.Write(HandleCMD(cmd));
        }

        // 采集控制 stop
        private void AcqControlStop(int MODE = 0, int RATE = 6, int START = 0)
        {
            string cmd = BCICommand.ACQUISTION_CONTROL;

            string mode = Convert.ToString(MODE, 2);
            mode = HandleLen(mode, 3);
            string rate = Convert.ToString(RATE, 2);
            rate = HandleLen(rate, 3);
            string start = Convert.ToString(START, 2);
            string byte1 = '0' + mode + rate + start;
            byte1 = "//x" + Convert.ToInt32(byte1, 2).ToString("X2");
            cmd = cmd + byte1;

            port.Write(HandleCMD(cmd));
        }

        // 心跳包
        private void HeartBeat()
        {
            port.Write(HandleCMD(BCICommand.HEART_BEAT));
        }

        // 查询寄存器
        private void QueryRegisiter()
        {
            port.Write(HandleCMD(BCICommand.QUERY_REGISTER));
        }

        // 软重启
        private void SoftRestart()
        {
            port.Write(HandleCMD(BCICommand.SOFT_RESTART));
            CmdNum = -1;
        }

        //查询版本
        private void QueryVersion()
        {
            port.Write(HandleCMD(BCICommand.QUERY_VERSION));
        }

        // 数据重发
        private void ResendData(int num)
        {
            string cmd = BCICommand.DATA_RESEND + num.ToString("X2");
            port.Write(HandleCMD(cmd));
        }

        #endregion

    }

    public struct BCICommand
    {
        // 以下命令不包含消息序号
        // 以下命令不包含校验和
        public const string CMD_FRAME_HEAD="\x55\xAA";
        public const string OPEN_CHANNEL = CMD_FRAME_HEAD + "\x00\x06" + "\x01" + "\x1F\xFF\xFF";  //开通道
        public const string CLOSE_CHANNEL = CMD_FRAME_HEAD + "\x00\x06" + "\x01" + "\x00\x00\x00";  //关通道
        public const string TEST_CHANNEL = CMD_FRAME_HEAD + "\x00\x05" + "\x03";  //测试通道 需要指定通道号和类型
        public const string SET_CHANNEL = CMD_FRAME_HEAD + "\x00\x05" + "\x05";  //通道设置 需要指定设置
        public const string TEST_IMPEDANCE = CMD_FRAME_HEAD + "\x00\x04" + "\x07";  //阻抗设置 需要指定命令内容
        public const string ACQUISTION_CONTROL = CMD_FRAME_HEAD + "\x00\x04" + "\x09";  //采集控制 需要指定命令内容
        public const string HEART_BEAT= CMD_FRAME_HEAD + "\x00\x03" + "\x0B"; //心跳包
        public const string QUERY_REGISTER= CMD_FRAME_HEAD + "\x00\x03" + "\x0C"; //寄存器
        public const string SOFT_RESTART= CMD_FRAME_HEAD + "\x00\x03" + "\x0F"; //软重启
        public const string QUERY_VERSION = CMD_FRAME_HEAD + "\x00\x03" + "\x11"; //查询版本
        public const string DATA_RESEND = CMD_FRAME_HEAD + "\x00\x04" + "\x13"; //查询版本

    }
}
