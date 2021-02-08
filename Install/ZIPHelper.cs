using ICSharpCode.SharpZipLib.Zip;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Install
{
    /// <summary>
    /// ZIP助手类
    /// </summary>
    public static class ZIPHelper
    {
        public static Action<double, double, string> ActionProgress;
        /// <summary>
        /// 解压缩zip文件
        /// </summary>
        /// <param name="zipFile">解压的zip文件流</param>
        /// <param name="extractPath">解压到的文件夹路径</param>
        /// <param name="bufferSize">读取文件的缓冲区大小</param>
        /// <param name="prog">进度条显示控件</param>
        /// <param name="tb">进度百分比显示控件</param>
        public static void Extract(byte[] zipFile, string extractPath, int bufferSize)
        {
            extractPath = extractPath.TrimEnd('/') + "//";
            byte[] data = new byte[bufferSize];
            int size;//缓冲区的大小（字节）
            double max = 0;//带待压文件的大小（字节）
            double osize = 0;//每次解压读取数据的大小（字节）
            using (ZipInputStream s = new ZipInputStream(new System.IO.MemoryStream(zipFile)))
            {
                ZipEntry entry;
                while ((entry = s.GetNextEntry()) != null)
                {
                    max += entry.Size;//获得待解压文件的大小
                }
            }
            using (ZipInputStream s = new ZipInputStream(new System.IO.MemoryStream(zipFile)))
            {
                ZipEntry entry;

                while ((entry = s.GetNextEntry()) != null)
                {
                    string directoryName = Path.GetDirectoryName(entry.Name);
                    string fileName = Path.GetFileName(entry.Name);

                    //先创建目录
                    if (directoryName.Length > 0)
                    {
                        Directory.CreateDirectory(extractPath + directoryName);
                    }
                    if (fileName != String.Empty)
                    {
                        using (FileStream streamWriter = File.Create(extractPath + entry.Name.Replace("/", "//")))
                        {
                            while (true)
                            {
                                size = s.Read(data, 0, data.Length);
                                if (size > 0)
                                {
                                    osize += size;
                                    System.Windows.Forms.Application.DoEvents();
                                    streamWriter.Write(data, 0, size);
                                    string text = Math.Round((osize / max * 100), 0).ToString() + "%";
                                    ActionProgress?.Invoke(max + 5, osize, text);

                                    System.Windows.Forms.Application.DoEvents();
                                }
                                else
                                {
                                    break;
                                }
                            }
                        }
                    }
                }
            }
        }
    }
}
