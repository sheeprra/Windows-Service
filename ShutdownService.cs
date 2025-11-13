//using System;
//using System.Diagnostics;
//using System.IO;
//using System.Reflection;
//using System.Runtime.InteropServices;
//using System.ServiceProcess;

//namespace AutoUPS
//{
//    public partial class MyShutdownService : ServiceBase
//    {
//        private SearchIO searchIO;

//        [StructLayout(LayoutKind.Sequential)]
//        public struct SERVICE_STATUS
//        {
//            public int dwServiceType;
//            public int dwCurrentState;
//            public int dwControlsAccepted;
//            public int dwWin32ExitCode;
//            public int dwServiceSpecificExitCode;
//            public int dwCheckPoint;
//            public int dwWaitHint;
//        }

//        [DllImport("advapi32.dll", SetLastError = true)]
//        private static extern bool SetServiceStatus(IntPtr handle, ref SERVICE_STATUS serviceStatus);

//        public MyShutdownService()
//        {
//            this.ServiceName = "MyShutdownService";
//            this.CanShutdown = true; //允许接收关机事件
//            this.CanStop = true;//是否在收到系统关机消息时触发 OnShutdown()
//                                //this.CanHandlePowerEvent = true;//是否处理电源事件（如挂起/恢复）
//                                //this.CanHandleSessionChangeEvent = true;//是否处理用户登录/注销事件


//            searchIO = new SearchIO();
//        }

//        private SERVICE_STATUS serviceStatus;

//        protected override void OnStart(string[] args)
//        {
//            serviceStatus.dwCurrentState = 0x00000004; // SERVICE_RUNNING
//            serviceStatus.dwControlsAccepted = 0x00000100 | 0x00000040;// 预关机 + 普通关机

//            SetServiceStatus(this.ServiceHandle, ref serviceStatus);

//            string startLog = GetLogWithTimestamp("服务启动成功");
//            File.AppendAllText("C:\\Service.log", startLog);

//            //延时5s
//            System.Threading.Thread.Sleep(5000);
//            searchIO.CacheTargetPortOnStartup();

//            AutoStart s = new AutoStart();
//            s.Ext();
//        }

//        protected override void OnShutdown()
//        {
//            string shutdownTriggerLog = GetLogWithTimestamp("系统触发关机事件，开始执行串口操作");
//            File.AppendAllText("C:\\Service.log", shutdownTriggerLog);
//            try
//            {
//                AutoClose p = new AutoClose();
//                p.Ext();
//                System.Threading.Thread.Sleep(1000);
//            }
//            catch (Exception ex)
//            {
//                string errorLog = GetLogWithTimestamp($"关机阶段遇到问题：{ex.Message} | 异常类型：{ex.GetType().Name}");
//                File.AppendAllText("C:\\Service.log", errorLog);
//            }
//            base.OnShutdown();
//        }

//        // 可选：如果你希望在 *更早的* 关机阶段执行（优先级更高）
//        protected override void OnCustomCommand(int command)
//        {
//            const int SERVICE_CONTROL_PRESHUTDOWN = 15;

//            string commandLog = GetLogWithTimestamp($"收到自定义命令：{command}");
//            File.AppendAllText("C:\\Service.log", commandLog);

//            if (command == SERVICE_CONTROL_PRESHUTDOWN)
//            {
//                File.AppendAllText("C:\\Service.log", "[PRESHUTDOWN] 收到系统预关机信号\r\n");

//                try
//                {
//                    AutoClose p = new AutoClose();
//                    p.Ext(); // 串口关闭动作在此
//                    System.Threading.Thread.Sleep(1000);
//                }
//                catch (Exception ex)
//                {
//                    string preshutdownErrorLog =
//                        GetLogWithTimestamp($"PRESHUTDOWN阶段执行失败：{ex.Message} | 堆栈信息：{ex.StackTrace.Substring(0, 200)}");
//                    File.AppendAllText("C:\\Service.log", preshutdownErrorLog);
//                }
//            }
//        }
//        private string GetLogWithTimestamp(string logContent)
//        {
//            // 时间格式：年-月-日 时:分:秒.毫秒（精确到毫秒，避免同秒日志混淆）
//            string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
//            // 拼接格式：[时间戳] | 日志内容 + 换行符（方便后续查看时按行分割）
//            return $"[{timestamp}] | {logContent}\r\n";
//        }
//    }
//}

using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.ServiceProcess;
using System.Threading;

namespace AutoUPS
{
    public partial class MyShutdownService : ServiceBase
    {
        private SearchIO searchIO;

        [StructLayout(LayoutKind.Sequential)]
        public struct SERVICE_STATUS
        {
            public int dwServiceType;              // 服务类型
            public int dwCurrentState;             // 当前状态（运行中、停止等）
            public int dwControlsAccepted;         // 可接受的控制命令
            public int dwWin32ExitCode;            // 错误码
            public int dwServiceSpecificExitCode;  // 特定服务错误码
            public int dwCheckPoint;               // 检查点（用于服务启动/停止进度）
            public int dwWaitHint;                 // 等待提示时间（ms）
        }



        [DllImport("advapi32.dll", SetLastError = true)]
        private static extern bool SetServiceStatus(IntPtr handle, ref SERVICE_STATUS serviceStatus);

        private SERVICE_STATUS serviceStatus;

        public MyShutdownService()
        {
            this.ServiceName = "MyShutdownService";
            this.CanShutdown = true;
            this.CanStop = true;
            this.CanHandlePowerEvent = false;
            this.CanHandleSessionChangeEvent = false;

            // 注册 PRESHUTDOWN 支持
            RegisterPreShutdownEvent();

            searchIO = new SearchIO();
        }

        /// <summary>
        /// 在服务构造阶段注册 PRESHUTDOWN 控制支持
        /// </summary>
        private void RegisterPreShutdownEvent()
        {
            FieldInfo field = typeof(ServiceBase).GetField("acceptedCommands", BindingFlags.Instance | BindingFlags.NonPublic);
            if (field == null)
            {
                throw new Exception("acceptedCommands field not found");
            }

            int value = (int)field.GetValue(this);
            field.SetValue(this, value | 256); // SERVICE_ACCEPT_PRESHUTDOWN (0x100)
        }

        protected override void OnStart(string[] args)
        {
            serviceStatus.dwServiceType = 0x10; // SERVICE_WIN32_OWN_PROCESS
            serviceStatus.dwCurrentState = 0x00000004; // SERVICE_RUNNING
            serviceStatus.dwControlsAccepted = 0x00000100 | 0x00000040; // PRESHUTDOWN + SHUTDOWN
            SetServiceStatus(this.ServiceHandle, ref serviceStatus);

            File.AppendAllText("C:\\Service.log", GetLog("服务启动成功"));

            // 启动时缓存串口
            Thread.Sleep(2000);
            searchIO.CacheTargetPortOnStartup();

            try
            {
                new AutoStart().Ext();
            }
            catch (Exception ex)
            {
                File.AppendAllText("C:\\Service.log", GetLog($"启动阶段异常：{ex.Message}"));
            }
        }

        /// <summary>
        /// SCM发送PRESHUTDOWN(15)时调用
        /// </summary>
        protected override void OnCustomCommand(int command)
        {
            base.OnCustomCommand(command);
            if (command == 15)
            {
                File.AppendAllText("C:\\Service.log", GetLog("[PRESHUTDOWN] 收到系统预关机信号"));
                this.OnPreShutdown();
            }
        }


        /// <summary>
        /// 预关机阶段自定义处理
        /// </summary>
        protected virtual void OnPreShutdown()
        {
            try
            {
                File.AppendAllText("C:\\Service.log", GetLog("执行预关机逻辑：关闭串口..."));
                AutoClose p = new AutoClose();
                p.Ext();
                Thread.Sleep(1000);
                File.AppendAllText("C:\\Service.log", GetLog("预关机逻辑执行完成。"));
            }
            catch (Exception ex)
            {
                File.AppendAllText("C:\\Service.log", GetLog($"[PRESHUTDOWN ERROR] {ex.Message}"));
            }
        }

        protected override void OnShutdown()
        {
            File.AppendAllText("C:\\Service.log", GetLog("[SHUTDOWN] 系统触发最终关机信号"));

            try
            {
                AutoClose p = new AutoClose();
                p.Ext();
                Thread.Sleep(1000);
            }
            catch (Exception ex)
            {
                File.AppendAllText("C:\\Service.log", GetLog($"[SHUTDOWN ERROR] {ex.Message}"));
            }

            base.OnShutdown();
        }

        
        private string GetLog(string content)
        {
            return $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] {content}\r\n";
        }
    }
}
