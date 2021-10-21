using Extensibility;
using Microsoft.Office.Core;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Windows.Forms;
using System.Xml.Linq;
using OneNote = Microsoft.Office.Interop.OneNote;


namespace litingaddin
{
    [Guid("308DEEE4-9577-431A-8940-3C1A8418BD06"), ProgId("litingaddin.Class1")]
    public class Class1 : IDTExtensibility2, IRibbonExtensibility
    {
        private OneNote.Application onApp = new OneNote.Application();
        private object application;
        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            application = Application;
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            application = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        public void OnAddInsUpdate(ref Array custom)
        {

        }

        public void OnStartupComplete(ref Array custom)
        {

        }

        public void OnBeginShutdown(ref Array custom)
        {
            if (application != null)
            {
                application = null;
            }
        }

        public string GetCustomUI(string RibbonID)
        {
            return Properties.Resources.ribbon;
        }

        public static void update_tittle_all()
        {
            OneNote.Application onenoteApp = new OneNote.Application();
            string xml;
            var pageid = onenoteApp.Windows.CurrentWindow.CurrentPageId;
            onenoteApp.GetPageContent(pageid, out xml, OneNote.PageInfo.piAll);
            var doc = XDocument.Parse(xml);
            XNamespace ns = doc.Root.Name.Namespace;
            var outLine_title = doc.Descendants(ns + "T").FirstOrDefault();
            //MessageBox.Show(outLine.Value);
            //XElement element =doc.Descendants(ns + "TagDef").FirstOrDefault();
            var title_update = string.Empty;
            int index = outLine_title.Value.LastIndexOf('｜');
            outLine_title.Value = outLine_title.Value.Substring(index + 1);
            foreach (XElement tags in from node in doc.Descendants(ns +
      "TagDef")
                                      select node)
            {
                var outLine_tag = tags.Attribute("name").Value.ToString();
                if (outLine_tag == "Page Tags")
                {
                    break;
                }
                else if(outLine_title.Value.Contains(outLine_tag)==true)
                {
                    break;
                }
                else if (outLine_title.Value.Contains(outLine_tag) == false)
                {
                    outLine_title.Value = outLine_tag + "｜" + outLine_title.Value;
                }
                else
                {
                    outLine_title.Value = outLine_tag + "｜" + outLine_title.Value;
                }
            }

            onenoteApp.UpdatePageContent(doc.ToString(), System.DateTime.MinValue);
        }
        public void update_tittle(IRibbonControl control)
        {
            update_tittle_all();
        }
        public static void Set_tags(IRibbonControl control,string p_type,string p_name)
        {
            OneNote.Application onenoteApp = new OneNote.Application();
            string xml;
            var pageid = onenoteApp.Windows.CurrentWindow.CurrentPageId;
            onenoteApp.GetPageContent(pageid, out xml, OneNote.PageInfo.piAll);
            var doc = XDocument.Parse(xml);
            XNamespace ns = doc.Root.Name.Namespace;
            string new_time = DateTime.Now.ToString("yyyy-MM-ddThh:mm:ss.fffZ");
            if (doc.Descendants(ns + "TagDef").Any() == true)
            {
                XElement TagDefs = doc.Descendants(ns + "TagDef").Last();
                //MessageBox.Show(TagDefs.ToString());
                string index = TagDefs.Attribute("index").Value;
                int index_i = int.Parse(index);
                index_i = index_i + 1;
                string index_s = index_i.ToString();
                XElement newTagDefs = new XElement(ns + "TagDef",
                                                new XAttribute("index", index_s),
                                                new XAttribute("type", p_type),
                                                new XAttribute("symbol", "0"),
                                                new XAttribute("fontColor", "automatic"),
                                                new XAttribute("highlightColor", "none"),
                                                new XAttribute("name", p_name)
                                                );
                //MessageBox.Show(newTagDefs.ToString());
                TagDefs.AddAfterSelf(newTagDefs);
                //TagDefs.Add(newTagDefs);

                XElement Tags = doc.Descendants(ns + "Tag").Last();
                XElement newTags = new XElement(ns + "Tag",
                                                new XAttribute("index", index_s),
                                                new XAttribute("completed", "true"),
                                                new XAttribute("disabled", "false"),
                                                new XAttribute("creationDate", new_time),
                                                new XAttribute("completionDate", new_time)
                                                );
                Tags.AddAfterSelf(newTags);
                //MessageBox.Show(doc.ToString());
                onenoteApp.UpdatePageContent(doc.ToString(), System.DateTime.MinValue);
            }
            else
            {
                XElement page = doc.Descendants(ns + "Page").FirstOrDefault();
                XElement newTagDefs = new XElement(ns + "TagDef",
                                                new XAttribute("index", "0"),
                                                new XAttribute("type", p_type),
                                                new XAttribute("symbol", "0"),
                                                new XAttribute("fontColor", "automatic"),
                                                new XAttribute("highlightColor", "none"),
                                                new XAttribute("name", p_name)
                                                );
                //MessageBox.Show(newTagDefs.ToString());
                page.AddFirst(newTagDefs);
                XElement OE = doc.Descendants(ns + "OE").FirstOrDefault();
                XElement newTags = new XElement(ns + "Tag",
                                                new XAttribute("index", "0"),
                                                new XAttribute("completed", "true"),
                                                new XAttribute("disabled", "false"),
                                                new XAttribute("creationDate", new_time),
                                                new XAttribute("completionDate", new_time)
                                                );
                OE.AddFirst(newTags);
                //MessageBox.Show(doc.ToString());
                onenoteApp.UpdatePageContent(doc.ToString(), System.DateTime.MinValue);
            }

            update_tittle_all();
        }
        public void Playlist_kaizhanzhong(IRibbonControl control)
        {
            Set_tags(control,"0", "【开展中】");
        }
        public void playlist_add(IRibbonControl control)
        {
            Set_tags(control, "0", "【未开展】");
        }
        public void Playlist_weiqueren(IRibbonControl control)
        {
            Set_tags(control, "2", "【未确认】");
        }

        public void Playlist_zuofei(IRibbonControl control)
        {
            Set_tags(control, "3", "【作废】");
        }
        public void Playlist_daisheji(IRibbonControl control)
        {
            Set_tags(control, "4", "【待设计】");
        }
        public void Playlist_weizhuan(IRibbonControl control)
        {
            Set_tags(control, "5", "【未转】");
        }
        public void Playlist_hebing(IRibbonControl control)
        {
            Set_tags(control, "6", "【合并】");
        }
        public void Playlist_yizhuan(IRibbonControl control)
        {
            Set_tags(control, "7", "【已转】");
        }

        public void Playlist_zanbukaizhan(IRibbonControl control)
        {
            Set_tags(control, "8", "【暂不开展】");
        }

        public void Playlist_yizhuanxubuchong(IRibbonControl control)
        {
            Set_tags(control, "9", "【已转需补充】");
        }

        public void zhengwen_duiqi()
        {

        }
        public void allin_xml(IRibbonControl control)
        {

            OneNote.Application onenoteApp = new OneNote.Application();
            string xml;
            var pageid = onenoteApp.Windows.CurrentWindow.CurrentPageId;
            onenoteApp.GetPageContent(pageid, out xml, OneNote.PageInfo.piAll);
            var doc = XDocument.Parse(xml);
            Application.EnableVisualStyles();
            //Application.SetCompatibleTextRenderingDefault(false);
            allinxml box1 = new allinxml();
            box1.textBox1.AppendText(doc.ToString());
            Application.Run(box1);
        }
        class CCOMStreamWrapper : IStream
        {
            public CCOMStreamWrapper(System.IO.Stream streamWrap)
            {
                m_stream = streamWrap;
            }

            public void Clone(out IStream ppstm)
            {
                ppstm = new CCOMStreamWrapper(m_stream);
            }

            public void Commit(int grfCommitFlags)
            {
                m_stream.Flush();
            }

            public void CopyTo(IStream pstm, long cb, IntPtr pcbRead, IntPtr pcbWritten)
            {
            }

            public void LockRegion(long libOffset, long cb, int dwLockType)
            {
                throw new System.NotImplementedException();
            }

            public void Read(byte[] pv, int cb, IntPtr pcbRead)
            {
                Marshal.WriteInt64(pcbRead, m_stream.Read(pv, 0, cb));
            }

            public void Revert()
            {
                throw new System.NotImplementedException();
            }

            public void Seek(long dlibMove, int dwOrigin, IntPtr plibNewPosition)
            {
                long posMoveTo = 0;
                Marshal.WriteInt64(plibNewPosition, m_stream.Position);
                switch (dwOrigin)
                {
                    case 0:
                        {
                            /* STREAM_SEEK_SET */
                            posMoveTo = dlibMove;
                        }
                        break;
                    case 1:
                        {
                            /* STREAM_SEEK_CUR */
                            posMoveTo = m_stream.Position + dlibMove;

                        }
                        break;
                    case 2:
                        {
                            /* STREAM_SEEK_END */
                            posMoveTo = m_stream.Length + dlibMove;
                        }
                        break;
                    default:
                        return;
                }
                if (posMoveTo >= 0 && posMoveTo < m_stream.Length)
                {
                    m_stream.Position = posMoveTo;
                    Marshal.WriteInt64(plibNewPosition, m_stream.Position);
                }
            }

            public void SetSize(long libNewSize)
            {
                m_stream.SetLength(libNewSize);
            }

            public void Stat(out System.Runtime.InteropServices.ComTypes.STATSTG pstatstg, int grfStatFlag)
            {
                pstatstg = new System.Runtime.InteropServices.ComTypes.STATSTG();
                pstatstg.cbSize = m_stream.Length;
                if ((grfStatFlag & 0x0001/* STATFLAG_NONAME */) != 0)
                    return;
                pstatstg.pwcsName = m_stream.ToString();
            }

            public void UnlockRegion(long libOffset, long cb, int dwLockType)
            {
                throw new System.NotImplementedException();
            }

            public void Write(byte[] pv, int cb, IntPtr pcbWritten)
            {
                Marshal.WriteInt64(pcbWritten, 0);
                m_stream.Write(pv, 0, cb);
                Marshal.WriteInt64(pcbWritten, cb);
            }

            private System.IO.Stream m_stream;
        }
        public IStream GetImage(string imageName)
        {
            MemoryStream mem = new MemoryStream();
            switch (imageName)
            {
                case "update_title.png":
                    Properties.Resources.update_title.Save(mem, ImageFormat.Png);
                    break;

                case "playlist_add.png":
                    Properties.Resources.playlist_add.Save(mem, ImageFormat.Png);
                    break;
                default:
                    Properties.Resources.update_title.Save(mem, ImageFormat.Png);
                    break;
            }

            return new CCOMStreamWrapper(mem);
        }



    }
}
