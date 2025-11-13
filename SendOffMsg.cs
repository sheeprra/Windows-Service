using System;
using System.IO;
using System.IO.Ports;
using System.Threading;

namespace AutoUPS
{
    /// <summary>
    /// CyberPower UPS 串口通信协议III实现（新增自定义指令发送）
    /// </summary>
    public class UPSCommunication : IDisposable
    {
        // 原有字段/构造函数/基础方法（保持不变，此处省略以简化代码，实际需保留）
        private SerialPort _serialPort;
        private const int TIMEOUT = 250;
        private readonly object _lockObj = new object();



        public UPSCommunication(string portName)
        {
            _serialPort = new SerialPort
            {
                PortName = portName,
                BaudRate = 2400,
                DataBits = 8,
                Parity = Parity.None,
                StopBits = StopBits.One,
                ReadTimeout = TIMEOUT,
                WriteTimeout = TIMEOUT,
                DtrEnable = true
            };

        }

        // 原有基础方法：Open/Close/CalculateCRC8/BuildPacket/SendPacket/ReceivePacket/ValidatePacket/SendACK/SendNAK/WaitForACK
        // （保持不变，此处省略，实际需保留原代码）
        /// <summary>
        /// CRC-8校验算法（原代码不变）
        /// </summary>
        private byte CalculateCRC8(byte[] data, int length)
        {
            const byte poly8 = 0xD5;
            byte crc = 0;
            for (int idx = 0; idx < length; idx++)
            {
                crc ^= data[idx];
                for (int i = 0; i < 8; i++)
                {
                    if ((crc & 0x80) != 0)
                        crc = (byte)((crc << 1) ^ poly8);
                    else
                        crc <<= 1;
                }
            }
            return crc;
        }

        /// <summary>
        /// 构建数据包
        /// </summary>
        private byte[] BuildPacket(byte[] payload)
        {
            int payloadSize = payload.Length;
            byte[] packet = new byte[payloadSize + 2]; // Head(1) + Payload(n) + Checksum(1)

            // Head字节：PayloadSize(6位) + Finish(1位=1) + Error(1位=0)
            byte head = (byte)(payloadSize & 0x3F);
            head |= 0x40; // Finish=1（单包传输）

            packet[0] = head;
            Array.Copy(payload, 0, packet, 1, payloadSize);

            // 计算CRC-8（Head+Payload）
            byte checksum = CalculateCRC8(packet, payloadSize + 1);
            packet[payloadSize + 1] = checksum;

            return packet;
        }

        /// <summary>
        /// 发送数据包
        /// </summary>
        private void SendPacket(byte[] packet)
        {
            _serialPort.Write(packet, 0, packet.Length);
            Thread.Sleep(10); // 等待传输完成（参考🔶1-59）
        }

        /// <summary>
        /// 接收数据包
        /// </summary>
        private byte[] ReceivePacket()
        {
            Thread.Sleep(10); // 等待传输结束（参考🔶1-59）
            int available = _serialPort.BytesToRead;
            if (available == 0)
                return null;

            byte[] buffer = new byte[available];
            _serialPort.Read(buffer, 0, available);
            return buffer;
        }

        /// <summary>
        /// 验证数据包
        /// </summary>
        private bool ValidatePacket(byte[] packet, out byte[] payload)
        {
            payload = null;
            if (packet == null || packet.Length < 2)
                return false;

            byte head = packet[0];
            int payloadSize = head & 0x3F;
            bool isError = (head & 0x80) != 0;

            if (isError || packet.Length != payloadSize + 2)
                return false;

            // 验证CRC-8
            byte receivedChecksum = packet[packet.Length - 1];
            byte calculatedChecksum = CalculateCRC8(packet, packet.Length - 1);
            if (receivedChecksum != calculatedChecksum)
                return false;

            payload = new byte[payloadSize];
            Array.Copy(packet, 1, payload, 0, payloadSize);
            return true;
        }

        /// <summary>
        /// 发送ACK/NAK
        /// </summary>
        private void SendACK() { _serialPort.Write(new byte[] { 0x40 }, 0, 1); Thread.Sleep(10); }
        private void SendNAK() { _serialPort.Write(new byte[] { 0xC0 }, 0, 1); Thread.Sleep(10); }

        /// <summary>
        /// 接收ACK/NAK
        /// </summary>
        private bool WaitForACK()
        {
            try
            {
                byte[] response = new byte[1];
                int bytesRead = _serialPort.Read(response, 0, 1);
                if (bytesRead == 1)
                {
                    return response[0] == 0x40; // 0x40=ACK，0xC0=NAK
                }
            }
            catch (TimeoutException) { }
            return false;
        }

        // ======================================
        // 发送自定义指令（0xC6 0x01 0x00）
        // ======================================
        /// <summary>
        /// 发送自定义指令（Payload：0xC6 0x01 0x00），对应完整封包：0x43 0xC6 0x01 0x00 0x0F
        /// </summary>
        /// <returns>UPS响应Payload（若成功），null（若失败）</returns>
        public byte[] SendCustomCommand_C60200(int maxRetries = 1)
        {
            lock (_lockObj)
            {
                // 1. 定义自定义指令的Payload（核心：0xC6 0x01 0x00）
                byte[] customPayload = new byte[] { 0xC6, 0x02, 0x00 };

                // 2. 构建完整数据包（调用原BuildPacket方法，自动生成Head和Checksum）
                // 预期生成的数据包：Head(0x43) + Payload(0xC6 0x01 0x00) + Checksum(0x0F)
                byte[] requestPacket = BuildPacket(customPayload);

                // 3. 重试机制（与原ExecuteCommand逻辑一致）
                for (int retry = 0; retry < maxRetries; retry++)
                {
                    try
                    {
                        // 步骤1：发送自定义指令包
                        SendPacket(requestPacket);
                        //Console.WriteLine($"已发送自定义指令包：{BitConverter.ToString(requestPacket)}"); // 验证是否为0x43-0xC6-0x01-0x00-0x0F

                        // 步骤2：等待UPS的ACK确认（参考🔶1-64）
                        if (!WaitForACK())
                        {
                            //Console.WriteLine("未收到ACK，重试...");
                            Thread.Sleep(50);
                            continue;
                        }
                        //Console.WriteLine("收到UPS的ACK确认");

                        // 步骤3：等待并接收UPS的响应包（参考🔶1-69，等待10-50ms）
                        Thread.Sleep(30);
                        byte[] responsePacket = ReceivePacket();
                        if (responsePacket == null)
                        {
                            //Console.WriteLine("未收到响应包，重试...");
                            continue;
                        }
                        //Console.WriteLine($"收到响应包：{BitConverter.ToString(responsePacket)}");

                        // 步骤4：验证响应包并返回Payload（参考🔶1-64）
                        if (ValidatePacket(responsePacket, out byte[] responsePayload))
                        {

                            SendACK(); // 向UPS发送ACK，确认接收响应
                            //Console.WriteLine("响应包验证成功，返回Payload");
                            return responsePayload;
                        }
                        else
                        {
                            SendNAK(); // 响应包无效，发送NAK
                            //Console.WriteLine("响应包验证失败，发送NAK");
                        }
                    }
                    catch (TimeoutException)
                    {
                        //Console.WriteLine("指令发送超时，清空缓冲区");
                        _serialPort.DiscardInBuffer();
                        _serialPort.DiscardOutBuffer();
                    }
                    catch (Exception ex)
                    {
                        //Console.WriteLine($"自定义指令异常：{ex.Message}");
                    }
                }

                //Console.WriteLine("自定义指令发送失败（已重试{maxRetries}次）");
                return null;
            }
        }

        public byte[] SendCustomCommand_C6FFFF(int maxRetries = 1)
        {
            lock (_lockObj)
            {
                // 1. 定义自定义指令的Payload（核心：0xC6 0x01 0x00）
                byte[] customPayload = new byte[] { 0xC6, 0xFF, 0xFF };

                // 2. 构建完整数据包（调用原BuildPacket方法，自动生成Head和Checksum）
                // 预期生成的数据包：Head(0x43) + Payload(0xC6 0x01 0x00) + Checksum(0x0F)
                byte[] requestPacket = BuildPacket(customPayload);

                // 3. 重试机制（与原ExecuteCommand逻辑一致）
                for (int retry = 0; retry < maxRetries; retry++)
                {
                    try
                    {
                        // 步骤1：发送自定义指令包
                        SendPacket(requestPacket);
                        //Console.WriteLine($"已发送自定义指令包：{BitConverter.ToString(requestPacket)}"); // 验证是否为0x43-0xC6-0x01-0x00-0x0F

                        // 步骤2：等待UPS的ACK确认（参考🔶1-64）
                        if (!WaitForACK())
                        {
                            //Console.WriteLine("未收到ACK，重试...");
                            Thread.Sleep(50);
                            continue;
                        }
                        //Console.WriteLine("收到UPS的ACK确认");

                        // 步骤3：等待并接收UPS的响应包（参考🔶1-69，等待10-50ms）
                        Thread.Sleep(30);
                        byte[] responsePacket = ReceivePacket();
                        if (responsePacket == null)
                        {
                            //Console.WriteLine("未收到响应包，重试...");
                            continue;
                        }
                        //Console.WriteLine($"收到响应包：{BitConverter.ToString(responsePacket)}");

                        // 步骤4：验证响应包并返回Payload（参考🔶1-64）
                        if (ValidatePacket(responsePacket, out byte[] responsePayload))
                        {
                            SendACK(); // 向UPS发送ACK，确认接收响应
                            return responsePayload;
                        }
                        else
                        {
                            SendNAK(); // 响应包无效，发送NAK
                        }
                    }
                    catch (TimeoutException)
                    {
                        _serialPort.DiscardInBuffer();
                        _serialPort.DiscardOutBuffer();
                    }
                    catch (Exception ex)
                    {
                      
                    }
                }

                return null;
            }
        }
         
        public void Dispose()
        {
            Close();
            _serialPort?.Dispose();
        }

        // 原有Open/Close方法（保持不变）
        public void Open() { if (!_serialPort.IsOpen) _serialPort.Open(); }
        public void Close() { if (_serialPort.IsOpen) _serialPort.Close(); }
    }

    /// <summary>
    /// 使用示例（新增自定义指令调用）
    /// </summary>
    class AutoClose
    {
        private SearchIO searchIO;
        public AutoClose()
        {
            searchIO = new SearchIO();
        }
        public void Ext()
        {
            File.AppendAllText("C:\\Service.log", GetLogWithTimestamp($"开始执行获取串口信息操作"));
            string port = searchIO.GetCachedTargetPort();
            if (port != null)
                port = port.Trim();
            using (var ups = new UPSCommunication(port)) // 替换为实际串口号
            {
                try
                {
                    ups.Open();
                    File.AppendAllText("C:\\Service.log", GetLogWithTimestamp($"打开串口"));
                    // 2. 新增：发送自定义指令（0xC6 0x01 0x00）
                    //Console.WriteLine("=== 发送自定义指令（0xC6 0x01 0x00）===");
                    byte[] customResponse = ups.SendCustomCommand_C60200(maxRetries: 1); // 执行1次
                    File.AppendAllText("C:\\Service.log", GetLogWithTimestamp($"发送串口信息完成"));
                    //if (customResponse != null)
                    //{

                    //}
                    //else
                    //{

                    //}
                    //关闭串口
                    ups.Close();
                }
                catch (Exception ex)
                {

                }
            }
        }
        static string GetLogWithTimestamp(string logContent)
        {
            string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
            return $"[{timestamp}] | {logContent}\r\n";
        }
    }

    class AutoStart
    {
        private SearchIO searchIO;
        public AutoStart()
        {
            searchIO = new SearchIO();
        }
        public void Ext()
        {
            File.AppendAllText("C:\\Service.log", GetLogWithTimestamp($"开始执行获取串口信息操作"));
            string port = searchIO.GetCachedTargetPort();
            if (port != null)
                port = port.Trim();
            using (var ups = new UPSCommunication(port)) // 替换为实际串口号
            {
                try
                {
                    ups.Open();
                    File.AppendAllText("C:\\Service.log", GetLogWithTimestamp($"打开串口"));
                    byte[] customResponse = ups.SendCustomCommand_C6FFFF(maxRetries: 1); // 执行1次
                    File.AppendAllText("C:\\Service.log", GetLogWithTimestamp($"发送串口信息完成"));
                    //if (customResponse != null)
                    //{

                    //}
                    //else
                    //{

                    //}
                    //关闭串口
                    ups.Close();
                }
                catch (Exception ex)
                {

                }
            }
        }
        static string GetLogWithTimestamp(string logContent)
        {
            string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
            return $"[{timestamp}] | {logContent}\r\n";
        }
    }

}


