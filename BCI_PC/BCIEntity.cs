using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Timers;
using System.Windows;

namespace BCI_PC
{
    class BCIEntity
    {
        public SerialPort port = new SerialPort();
        private List<byte> rawBuffer=new List<byte>();
        private int CmdNum = -1;  // 消息序号
        int notHeart = 0;
        bool isBadData = false; // 收到校验失败的数据
        bool isReceivedResendData = false; //收到重发数据
        int recentNo = -1;
        public bool isSave = true;

        public BCIEntity(string name, int rate = 921600)
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
            BCICommandTime.init();
            port.Open();
            port.DataReceived += Port_DataReceived;

            Timer t = new Timer(1000);
            t.Elapsed += new ElapsedEventHandler((object sender, ElapsedEventArgs e) => {
                HeartBeat();
                if (BCICommandTime.HEART_BEAT_Time)
                {
                    BCICommandTime.HEART_BEAT_Time = false;
                    notHeart = 0;
                }
                else
                {
                    notHeart++;
                    if (notHeart == 2)
                    {
                        MessageBox.Show("断开连接");
                        t.Stop();
                        if (port.IsOpen)
                        {
                            port.Close();
                        }
                    }
                }
            });
            t.AutoReset = true;
            t.Start();
        }

        #region 接收数据

        private bool CheckSum(byte[] response)
        {
            int len = response.Length;
            int sum = 0;
            for (int i = 2; i < len -1; i++)
            {
                sum += (int)response[i];
            }
            sum = sum % 256;
            if (sum == response[len - 1])
                return true;
            else
                return false;
        }

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
            if (isReceivedResendData)
            {

                HandleBadData();
            }

            DataParser();

        }

        private void HandleBadData()
        {
            int len = rawBuffer.Count;
            bool has = false;
            int start = 0;
            for(int i = 0; i < len; i++)
            {
                if(rawBuffer[i]==0xFF&& rawBuffer[i+1] == 0xFF&& rawBuffer[i + 5] == recentNo)
                {
                    if (len - i >= 76)
                    {
                        has = true;
                        start = i;
                        break;
                    }
                }
            }
            if (has)
            {
                // 将重发的数据插入头部
                var buf = new byte[76];
                rawBuffer.CopyTo(start,buf,0,76);
                rawBuffer.RemoveRange(start, 76);
                rawBuffer.InsertRange(0, buf);
            }
        }

        private void DataParser()
        {
            while (rawBuffer.Count >= 5)
            {
                if (rawBuffer[0] == 0x55 && rawBuffer[1] == 0xAA) //命令
                {
                    int frameLength = rawBuffer[2] << 8 | rawBuffer[3];
                    frameLength = frameLength + 2 + 2; // 帧头2byte 长度2byte
                    byte[] bodybuf = new byte[frameLength];
                    rawBuffer.CopyTo(0,bodybuf, 0, frameLength);
                    rawBuffer.RemoveRange(0, frameLength);
                    switch (rawBuffer[4])
                    {
                        case 0x02:  // 通道开控制
                            if (!CheckSum(bodybuf))
                            {
                                MessageBox.Show("开通道校验出错啦");
                                break;
                            }
                            BCICommandTime.OPEN_CHANNEL_Time = true;
                            if (bodybuf[6] == 0x01)
                            {
                                MessageBox.Show("开通道出错");
                            }
                            break;
                        case 0x04:  // 测试通道
                            if (!CheckSum(bodybuf))
                            {
                                MessageBox.Show("测试通道校验出错啦");
                                break;
                            }
                            BCICommandTime.TEST_CHANNEL_Time = true;
                            if (bodybuf[6] == 0x01)
                            {
                                MessageBox.Show("测试通道出错");
                            }
                            break;
                        case 0x06:  // 设置通道
                            if (!CheckSum(bodybuf))
                            {
                                MessageBox.Show("设置通道校验出错啦");
                                break;
                            }
                            BCICommandTime.SET_CHANNEL_Time = true;
                            if (bodybuf[6] == 0x01)
                            {
                                MessageBox.Show("设置通道出错");
                            }
                            break;
                        case 0x08:  // 阻抗设置
                            if (!CheckSum(bodybuf))
                            {
                                MessageBox.Show("阻抗设置校验出错啦");
                                break;
                            }
                            BCICommandTime.TEST_IMPEDANCE_time = true;
                            if (bodybuf[6] == 0x01)
                            {
                                MessageBox.Show("阻抗设置出错");
                            }
                            break;
                        case 0x0A:  // 采集控制
                            if (!CheckSum(bodybuf))
                            {
                                MessageBox.Show("采集控制校验出错啦");
                                break;
                            }
                            BCICommandTime.ACQUISTION_CONTROL_time = true;
                            if (bodybuf[6] == 0x01)
                            {
                                MessageBox.Show("采集控制出错");
                            }
                            break;
                        case 0x0C:  // 心跳包
                            if (!CheckSum(bodybuf))
                            {
                                MessageBox.Show("心跳包校验出错啦");
                                break;
                            }
                            BCICommandTime.HEART_BEAT_Time = true;
                            break;
                        case 0x0E:  // 寄存器
                            if (!CheckSum(bodybuf))
                            {
                                MessageBox.Show("寄存器校验出错啦");
                                break;
                            }
                            BCICommandTime.QUERY_REGISTER_Time = true;
                            break;
                        case 0x10:  // 软重启
                            if (!CheckSum(bodybuf))
                            {
                                MessageBox.Show("软重启校验出错啦");
                                break;
                            }
                            BCICommandTime.SOFT_RESTART_Time = true;
                            break;
                        case 0x12:  // 查询版本
                            if (!CheckSum(bodybuf))
                            {
                                MessageBox.Show("查询版本校验出错啦");
                                break;
                            }
                            BCICommandTime.QUERY_VERSION_Time = true;
                            break;
                        case 0x14:  // 数据重发
                            if (!CheckSum(bodybuf))
                            {
                                MessageBox.Show("数据重发校验出错啦");
                                break;
                            }
                            BCICommandTime.DATA_RESEND_Time = true;
                            if (bodybuf[6] == 0x01)
                            {
                                MessageBox.Show("重发数据出错");
                                break;
                            }
                            if (isBadData)
                            {
                                isReceivedResendData = true;
                            }
                            isBadData = false;
                            break;
                    }


                }
                else if (rawBuffer[0] == 0xFF && rawBuffer[1] == 0xFF && !isBadData) // 数据
                {
                    int frameLength = rawBuffer[2] << 8 | rawBuffer[3];
                    frameLength = frameLength + 2 + 2; // 帧头2byte 长度2byte
                    byte[] bodybuf = new byte[frameLength];
                    rawBuffer.CopyTo(0, bodybuf, 0, frameLength);
                    rawBuffer.RemoveRange(0, frameLength);
                    StringBuilder sbchannel = new StringBuilder();
                    StringBuilder sbxyz = new StringBuilder();
                    if (!CheckSum(bodybuf))
                    {
                        MessageBox.Show("数据接收校验出错啦");
                        isBadData = true;
                        ResendData(bodybuf[5]);
                        recentNo = bodybuf[5];
                        break;
                    }


                    // eeg
                    int[] channel = new int[21];
                    sbchannel.Append(DateTime.Now.ToString("HH:mmss")); sbchannel.Append(",");
                    for (int i = 0; i < 21; i++)
                    {
                        channel[20 - i] = bodybuf[6 + i * 3] << 16 | bodybuf[7 + i * 3] << 8 | bodybuf[8 + i * 3];
                        sbchannel.Append(channel[20 - i]);
                        if(i!=20)
                            sbchannel.Append(',');
                    }
                    sbchannel.AppendLine();
                    if (isSave)
                    {
                        SaveData(ref sbchannel, "EEG");
                    }

                    // save??



                    // 辅助数据
                    switch (bodybuf[4])
                    {
                        case 0x00:  //加速度传感器
                            int z = bodybuf[69] << 8 | bodybuf[70];
                            int y = bodybuf[71] << 8 | bodybuf[72];
                            int x = bodybuf[73] << 8 | bodybuf[74];
                            sbxyz.Append(DateTime.Now.ToString("HH:mmss")); sbxyz.Append(",");
                            sbxyz.Append(x); sbxyz.Append(",");
                            sbxyz.Append(y); sbxyz.Append(",");
                            sbxyz.Append(z); sbxyz.AppendLine();
                            if (isSave)
                            {
                                SaveData(ref sbxyz, "Xyz");
                            }
                            break;
                        case 0x01:  // 电量模式
                            int batt = bodybuf[74];
                            break;
                        case 0x02:  // 模拟模式
                            break;
                        case 0x03:  // 数字IO模式
                            break;
                        case 0x04: // 标记模式
                            int mark = bodybuf[74];
                            break;
                        case 0x05:  // 点击脱落检测模式
                            int byte1 = bodybuf[72];
                            int byte2 = bodybuf[73];
                            int byte3 = bodybuf[74];
                            int[] allstate = new int[21];  // 每个电极脱落与否
                            int stateI = 21;
                            string state = Convert.ToString(byte1, 2);
                            state = HandleLen(state, 5);
                            for(int i = 0; i < 5; i++)
                            {
                                stateI--;
                                allstate[stateI] = state[i];
                            }

                            state = Convert.ToString(byte2, 2);
                            state = HandleLen(state, 8);
                            for (int i = 0; i < 8; i++)
                            {
                                stateI--;
                                allstate[stateI] = state[i];
                            }

                            state = Convert.ToString(byte3, 2);
                            state = HandleLen(state, 8);
                            for (int i = 0; i < 8; i++)
                            {
                                stateI--;
                                allstate[stateI] = state[i];
                            }

                            break;
                    }
                }
                else
                {
                    rawBuffer.RemoveAt(0);
                }
            }

        }

        #endregion

        #region 处理命令
        // 添加校验和
        private string CalCheckSum(string AnteriorSegment)
        {
            int len = AnteriorSegment.Length;
            int sum = 0;
            for(int i = 2; i< len; i++)
            {
                sum += (int)AnteriorSegment[i];
            }
            sum = sum % 256;
            return AnteriorSegment+ Convert.ToString(len, 16);
        }

        // 获得消息序号
        private string GetCmdNum()
        {
            CmdNum++;
            if (CmdNum == 255)
            {
                CmdNum = 0;
            }
            return Convert.ToString(CmdNum, 16);
        }

        // 给命令添加消息序号
        private string AddCmdNum(string AnteriorSegment)
        {
            int pos = 2 + 2 + 2;
            AnteriorSegment.Insert(pos, GetCmdNum());
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
            Timer t = new Timer(200);
            t.Elapsed += new ElapsedEventHandler((object sender, ElapsedEventArgs e)=>{
                if (BCICommandTime.OPEN_CHANNEL_Time)
                {
                    BCICommandTime.OPEN_CHANNEL_Time = false;
                    BCICommandTime.resendCount[0] = 0;
                }
                else
                {
                    if (BCICommandTime.resendCount[0] < 3)
                    {
                        OpenChannel();
                    }
                    BCICommandTime.resendCount[0]++;
                }
            });
            t.AutoReset = false;
            t.Start();
        }

        // 通道关
        private void CloseChannel()
        {
            port.Write(HandleCMD(BCICommand.CLOSE_CHANNEL));
            Timer t = new Timer(200);
            t.Elapsed += new ElapsedEventHandler((object sender, ElapsedEventArgs e) => {
                if (BCICommandTime.OPEN_CHANNEL_Time)
                {
                    BCICommandTime.OPEN_CHANNEL_Time = false;
                    BCICommandTime.resendCount[0] = 0;
                }
                else
                {
                    if (BCICommandTime.resendCount[0] < 3)
                    {
                        OpenChannel();
                    }
                    BCICommandTime.resendCount[0]++;
                }
            });
            t.AutoReset = false;
            t.Start();
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
            string cmd = BCICommand.TEST_CHANNEL + Convert.ToString(channel, 16) + Convert.ToString(channelType, 16);
            port.Write(HandleCMD(cmd));

            Timer t = new Timer(200);
            t.Elapsed += new ElapsedEventHandler((object sender, ElapsedEventArgs e) => {
                if (BCICommandTime.TEST_CHANNEL_Time)
                {
                    BCICommandTime.TEST_CHANNEL_Time = false;
                    BCICommandTime.resendCount[1] = 0;
                }
                else
                {
                    if (BCICommandTime.resendCount[1] < 3)
                    {
                        TestChannel(channel,channelType);
                    }
                    BCICommandTime.resendCount[1]++;
                }
            });
            t.AutoReset = false;
            t.Start();
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
            byte1= Convert.ToString(Convert.ToInt32(byte1, 2), 16);

            string bias = Convert.ToString(BIAS, 2);
            string type = Convert.ToString(TYPE, 2);
            type = HandleLen(type, 3);
            string gan = Convert.ToString(GAN, 2);
            gan = HandleLen(gan, 3);
            string batt = Convert.ToString(BATT, 2);
            string byte2 = bias + type + gan+batt;
            byte2 = Convert.ToString(Convert.ToInt32(byte2, 2), 16);
            cmd = cmd + byte1 + byte2;

            port.Write(HandleCMD(cmd));


            Timer t = new Timer(200);
            t.Elapsed += new ElapsedEventHandler((object sender, ElapsedEventArgs e) => {
                if (BCICommandTime.SET_CHANNEL_Time)
                {
                    BCICommandTime.SET_CHANNEL_Time = false;
                    BCICommandTime.resendCount[2] = 0;
                }
                else
                {
                    if (BCICommandTime.resendCount[2] < 3)
                    {
                        SetChannel(CHANNEL, BATT, SRB1, SRB2, BIAS, TYPE, GAN);
                    }
                    BCICommandTime.resendCount[2]++;
                }
            });
            t.AutoReset = false;
            t.Start();
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
            byte1 = Convert.ToString(Convert.ToInt32(byte1, 2), 16);

            cmd = cmd + byte1;

            port.Write(HandleCMD(cmd));

            Timer t = new Timer(200);
            t.Elapsed += new ElapsedEventHandler((object sender, ElapsedEventArgs e) => {
                if (BCICommandTime.TEST_IMPEDANCE_time)
                {
                    BCICommandTime.TEST_IMPEDANCE_time = false;
                    BCICommandTime.resendCount[3] = 0;
                }
                else
                {
                    if (BCICommandTime.resendCount[3] < 3)
                    {
                        TestImpedance(CHANNEL, REVERSAL, N, P);
                    }
                    BCICommandTime.resendCount[3]++;
                }
            });
            t.AutoReset = false;
            t.Start();
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
            byte1 = Convert.ToString(Convert.ToInt32(byte1, 2), 16);
            cmd = cmd + byte1;

            port.Write(HandleCMD(cmd));


            Timer t = new Timer(200);
            t.Elapsed += new ElapsedEventHandler((object sender, ElapsedEventArgs e) => {
                if (BCICommandTime.ACQUISTION_CONTROL_time)
                {
                    BCICommandTime.ACQUISTION_CONTROL_time = false;
                    BCICommandTime.resendCount[4] = 0;
                }
                else
                {
                    if (BCICommandTime.resendCount[4] < 3)
                    {
                        AcqControlStart(MODE, RATE, START);
                    }
                    BCICommandTime.resendCount[4]++;
                }
            });
            t.AutoReset = false;
            t.Start();
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
            byte1 = Convert.ToString(Convert.ToInt32(byte1, 2), 16);
            cmd = cmd + byte1;

            port.Write(HandleCMD(cmd));

            Timer t = new Timer(200);
            t.Elapsed += new ElapsedEventHandler((object sender, ElapsedEventArgs e) => {
                if (BCICommandTime.ACQUISTION_CONTROL_time)
                {
                    BCICommandTime.ACQUISTION_CONTROL_time = false;
                    BCICommandTime.resendCount[4] = 0;
                }
                else
                {
                    if (BCICommandTime.resendCount[4] < 3)
                    {
                        AcqControlStart(MODE, RATE, START);
                    }
                    BCICommandTime.resendCount[4]++;
                }
            });
            t.AutoReset = false;
            t.Start();
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


            Timer t = new Timer(200);
            t.Elapsed += new ElapsedEventHandler((object sender, ElapsedEventArgs e) => {
                if (BCICommandTime.QUERY_REGISTER_Time)
                {
                    BCICommandTime.QUERY_REGISTER_Time = false;
                    BCICommandTime.resendCount[6] = 0;
                }
                else
                {
                    if (BCICommandTime.resendCount[6] < 3)
                    {
                        QueryRegisiter();
                    }
                    BCICommandTime.resendCount[6]++;
                }
            });
            t.AutoReset = false;
            t.Start();
        }

        // 软重启
        private void SoftRestart()
        {
            port.Write(HandleCMD(BCICommand.SOFT_RESTART));

            Timer t = new Timer(200);
            t.Elapsed += new ElapsedEventHandler((object sender, ElapsedEventArgs e) => {
                if (BCICommandTime.SOFT_RESTART_Time)
                {
                    BCICommandTime.SOFT_RESTART_Time = false;
                    BCICommandTime.resendCount[7] = 0;
                    CmdNum = -1;
                    BCICommandTime.init();
                }
                else
                {
                    if (BCICommandTime.resendCount[7] < 3)
                    {
                        SoftRestart();
                    }
                    BCICommandTime.resendCount[7]++;
                }
            });
            t.AutoReset = false;
            t.Start();
        }

        //查询版本
        private void QueryVersion()
        {
            port.Write(HandleCMD(BCICommand.QUERY_VERSION));

            Timer t = new Timer(200);
            t.Elapsed += new ElapsedEventHandler((object sender, ElapsedEventArgs e) => {
                if (BCICommandTime.QUERY_VERSION_Time)
                {
                    BCICommandTime.QUERY_VERSION_Time = false;
                    BCICommandTime.resendCount[8] = 0;
                }
                else
                {
                    if (BCICommandTime.resendCount[8] < 3)
                    {
                        QueryVersion();
                    }
                    BCICommandTime.resendCount[8]++;
                }
            });
            t.AutoReset = false;
            t.Start();
        }

        // 数据重发
        private void ResendData(int num)
        {
            string cmd = BCICommand.DATA_RESEND + num.ToString("X2");
            port.Write(HandleCMD(cmd));

            Timer t = new Timer(200);
            t.Elapsed += new ElapsedEventHandler((object sender, ElapsedEventArgs e) => {
                if (BCICommandTime.DATA_RESEND_Time)
                {
                    BCICommandTime.DATA_RESEND_Time = false;
                    BCICommandTime.resendCount[9] = 0;
                }
                else
                {
                    if (BCICommandTime.resendCount[9] < 3)
                    {
                        ResendData(num);
                    }
                    BCICommandTime.resendCount[9]++;
                }
            });
            t.AutoReset = false;
            t.Start();
        }

        #endregion

        #region save data

        public static string savePath = "";
        public static void initSave()
        {
            List<Tuple<string, string>> files = new List<Tuple<string, string>> // filename, file headers
            {
                new Tuple<string, string>(DateTime.Now.ToString("yyyy-MM-dd") + "_" +"Xyz", "Time, x,y,z"),
                new Tuple<string, string>(DateTime.Now.ToString("yyyy-MM-dd") + "_" +"EEG", "Time, 1, 2, 3, 4, 5, 6, 7, 8,9,10,11,12,13,14,15,16,17,18,19,20,21"),
            };
            foreach (var file in files)
            {
                var filename = savePath + @"\" + file.Item1 + @".csv";
                if (!File.Exists(filename))
                {
                    try
                    {
                        File.Create(filename).Close();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error creating record file {file}.csv: {ex.Message}", "Error");
                        return;
                    }
                    try
                    {
                        using (StreamWriter sw = new StreamWriter(filename, false))    // overwrite file, because it's newly created.
                        {
                            sw.Write(file.Item2);   // write header
                            sw.WriteLine();
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error writing record header to {file}.csv: {ex.Message}", "Error");
                    }
                }
            }
        }

        private void SaveData(ref StringBuilder sb, string kind)
        {
            try
            {
                using (StreamWriter sw = new StreamWriter(Path.Combine(savePath, DateTime.Now.ToString("yyyy-MM-dd") + "_" + kind + ".csv"), true))
                {
                    sw.Write(sb.ToString());
                    sb.Clear();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error saving to {kind}.csv: {ex.Message}", "Error");
            }
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

    public class BCICommandTime
    {
        // 10种命令
        public static bool OPEN_CHANNEL_Time = false;
        public static bool TEST_CHANNEL_Time = false;
        public static bool SET_CHANNEL_Time = false;
        public static bool TEST_IMPEDANCE_time = false;
        public static bool ACQUISTION_CONTROL_time = false;
        public static bool HEART_BEAT_Time = false;
        public static bool QUERY_REGISTER_Time = false;
        public static bool SOFT_RESTART_Time = false;
        public static bool QUERY_VERSION_Time = false;
        public static bool DATA_RESEND_Time = false;

        public static int[] resendCount = new int[10];



        public static void init()
        {
            OPEN_CHANNEL_Time = false;
            TEST_CHANNEL_Time = false;
            SET_CHANNEL_Time = false;
            TEST_IMPEDANCE_time = false;
            ACQUISTION_CONTROL_time = false;
            HEART_BEAT_Time = false;
            QUERY_REGISTER_Time = false;
            SOFT_RESTART_Time = false;
            QUERY_VERSION_Time = false;
            DATA_RESEND_Time = false;

            for (int i = 0; i < 10; i++)
            {
                resendCount[i] = 0;
            }
        }

    }
}
