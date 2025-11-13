using System;
using System.Diagnostics;
using System.IO.Ports;
using System.Management;
using System.Text.RegularExpressions;
using System.IO;

namespace AutoUPS
{
    internal class SearchIO
    {
            // 静态字段：缓存目标串口号（开机时赋值，全局复用）
        private static string _cachedTargetPort;
        private const int MaxRetryCount = 3; // 最多重试3次
        private const int RetryIntervalMs = 2000; // 每次重试间隔3秒（3000毫秒）

        // 🔴 开机时调用：查询并缓存目标串口（仅执行一次）
        // 🔴 开机时调用：查询并缓存目标串口（支持重试）
        public void CacheTargetPortOnStartup()
        {
            string targetKeyword = "Prolific";
            int retryCount = 0;

            // 循环重试：最多MaxRetryCount次
            while (retryCount < MaxRetryCount)
            {
                try
                {
                    string[] portNames = SerialPort.GetPortNames();
                    retryCount++; // 重试次数+1（首次执行也计为第1次）

                    // 记录当前重试次数和枚举到的串口
                    File.AppendAllText("C:\\Service.log", GetLogWithTimestamp($"【第{retryCount}/{MaxRetryCount}次查询】开机枚举串口：{string.Join(", ", portNames)}"));

                    // 遍历查询，找到目标串口并缓存
                    foreach (string portName in portNames)
                    {
                        string desc = GetPortDescription(portName);
                        if (!string.IsNullOrEmpty(desc) && desc.IndexOf(targetKeyword, StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            _cachedTargetPort = portName; // 缓存串口号
                            File.AppendAllText("C:\\Service.log", GetLogWithTimestamp($"【第{retryCount}次查询成功】开机缓存目标串口：{_cachedTargetPort}（描述：{desc}）"));
                            return; // 找到后直接返回，终止重试
                        }
                    }

                    // 走到这里说明本次查询未找到目标串口
                    if (retryCount < MaxRetryCount)
                    {
                        // 非最后一次重试：记录日志并暂停3秒
                        File.AppendAllText("C:\\Service.log", GetLogWithTimestamp($"【第{retryCount}次查询失败】未找到含关键字「{targetKeyword}」的串口，{RetryIntervalMs / 1000}秒后重试..."));
                        System.Threading.Thread.Sleep(RetryIntervalMs); // 暂停3秒
                    }
                    else
                    {
                        // 最后一次重试失败：记录日志并设缓存为null
                        File.AppendAllText("C:\\Service.log", GetLogWithTimestamp($"【第{retryCount}次查询失败】已达最大重试次数，未找到目标串口"));
                        _cachedTargetPort = null;
                    }
                }
                catch (Exception ex)
                {
                    // 异常处理：记录错误日志，若未达最大重试次数则继续，否则终止
                    File.AppendAllText("C:\\Service.log", GetLogWithTimestamp($"【第{retryCount}次查询异常】{ex.Message}"));
                    if (retryCount < MaxRetryCount)
                    {
                        File.AppendAllText("C:\\Service.log", GetLogWithTimestamp($"{RetryIntervalMs / 1000}秒后重试..."));
                        System.Threading.Thread.Sleep(RetryIntervalMs);
                    }
                    else
                    {
                        _cachedTargetPort = null;
                    }
                }
            }
        }

        // 🔴 关机时调用：直接获取缓存的串口号
        public string GetCachedTargetPort()
        {
            if (!string.IsNullOrEmpty(_cachedTargetPort))
            {
                File.AppendAllText("C:\\Service.log", GetLogWithTimestamp($"使用缓存串口：{_cachedTargetPort}"));
                return _cachedTargetPort;
            }

            // 极端情况：缓存为空（如开机时未找到），尝试最后一次快速查询（可选）
            File.AppendAllText("C:\\Service.log", GetLogWithTimestamp("缓存串口为空，尝试最后一次快速查询"));
            return QuickSearchTargetPort();
        }

        // 辅助方法：快速查询
        private string QuickSearchTargetPort()
        {
            return null;
        }

        // 修改：将 C# 8.0 的 using 声明替换为 C# 7.3 兼容的 using 语句  
        static string GetPortDescription(string portName)
        {
            try
            {
                using (var searcher = new ManagementObjectSearcher(
                    "SELECT * FROM Win32_PnPEntity WHERE Name LIKE '%(COM%'"))
                {
                    foreach (ManagementObject obj in searcher.Get())
                    {
                        string name = obj["Name"]?.ToString();
                        if (name != null && name.Contains(portName))
                            return name;
                    }
                }
            }
            catch (Exception ex)
            {
                // 开机时记录WMI查询异常（调试用）  
                File.AppendAllText("C:\\Service.log", GetLogWithTimestamp($"WMI查询异常：{ex.Message}"));
            }
            return "";
        }


        // 原有方法：生成带时间戳的日志
        static string GetLogWithTimestamp(string logContent)
        {
            string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
            return $"[{timestamp}] | {logContent}\r\n";
        }
    }
    }

