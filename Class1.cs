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
using System.Threading;
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
            //var outLine_title = doc.Descendants(ns + "T").FirstOrDefault();
            //MessageBox.Show(outLine.Value);
            //XElement element =doc.Descendants(ns + "TagDef").FirstOrDefault();
            XElement OutLine_title = doc.Descendants(ns + "Title").FirstOrDefault();
            foreach (XElement outLine_titles_all in from node in OutLine_title.Descendants(ns + "T") select node)
            {
                string outLine_titles_one = outLine_titles_all.Value;
                if (string.IsNullOrEmpty(outLine_titles_one))
                {
                    break;
                }
                else
                {
                    var title_update = string.Empty;
                    int index = outLine_titles_all.Value.LastIndexOf('｜');
                    outLine_titles_all.Value = outLine_titles_all.Value.Substring(index + 1);
                    foreach (XElement tags in from node in doc.Descendants(ns +
              "TagDef")
                                              select node)
                    {
                        var outLine_tag = tags.Attribute("name").Value.ToString();
                        if (outLine_tag == "Page Tags")
                        {
                            break;
                        }
                        else if (outLine_titles_all.Value.Contains(outLine_tag) == true)
                        {
                            break;
                        }
                        else if (outLine_titles_all.Value.Contains(outLine_tag) == false)
                        {
                            outLine_titles_all.Value = outLine_tag + "｜" + outLine_titles_all.Value;
                        }
                        else
                        {
                            outLine_titles_all.Value = outLine_tag + "｜" + outLine_titles_all.Value;
                        }
                    }
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
            Set_tags(control, "1", "【未开展】");
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
        public class OutLines_del
        {
            public string OutLines_del_data;
        }
        public void Del_none(IRibbonControl control)
        {
            OneNote.Application onenoteApp = new OneNote.Application();
            string xml;
            var pageid = onenoteApp.Windows.CurrentWindow.CurrentPageId;
            onenoteApp.GetPageContent(pageid, out xml, OneNote.PageInfo.piAll);
            var doc = XDocument.Parse(xml);
            XNamespace ns = doc.Root.Name.Namespace;
            //MessageBox.Show(TagDefs.ToString());
            string new_time = DateTime.Now.ToString("yyyy-MM-ddThh:mm:ss.fffZ");
            var OutLines_dels = new List<OutLines_del>();
            foreach (XElement Outlines in from node in doc.Descendants(ns + "Outline").ToList() select node)
            {
                string OutLine_data = Outlines.Descendants(ns + "T").FirstOrDefault().Value.ToString();
                int OutLine_count = Outlines.Descendants(ns + "T").Count();
                if (String.IsNullOrEmpty(OutLine_data) && (OutLine_count == 1) )
                {
                    string ObjectIDs = Outlines.Attribute("objectID").Value;
                    OutLines_dels.Add(new OutLines_del() { OutLines_del_data = ObjectIDs });
                }
                else
                {
                    continue;
                }
            }
            foreach (var OutLines_del_datas in OutLines_dels.ToList())
            {
                onenoteApp.DeletePageContent(pageid, OutLines_del_datas.OutLines_del_data.ToString(), System.DateTime.MinValue);
            }
        }

        public void Tongyi_data(IRibbonControl control)
        {
            OneNote.Application onenoteApp = new OneNote.Application();
            string xml;
            var pageid = onenoteApp.Windows.CurrentWindow.CurrentPageId;
            onenoteApp.GetPageContent(pageid, out xml, OneNote.PageInfo.piAll);
            var doc = XDocument.Parse(xml);
            XNamespace ns = doc.Root.Name.Namespace;
            foreach (XElement Outlines in from node in doc.Descendants(ns + "Outline") select node)
            {
                string OutLine_data = Outlines.Descendants(ns + "T").FirstOrDefault().Value.ToString();
                if (String.IsNullOrEmpty(OutLine_data))
                {
                    break;
                }
                else
                {
                    XElement Positions = Outlines.Descendants(ns + "Position").FirstOrDefault();
                    Positions.Attribute("x").Value= "36.00000000000000";
                    Positions.Attribute("y").Value= "103.30000000000000";
                    XElement Sizes = Outlines.Descendants(ns + "Size").FirstOrDefault();
                    string Size_w = Sizes.Attribute("width").Value;
                    string Size_h = Sizes.Attribute("height").Value;
                    double Size_w_int = double.Parse(Size_w);
                    double Size_h_int = double.Parse(Size_h);
                    double Size_chu = Size_h_int / Size_w_int;
                    Sizes.Attribute("width").Value = "778.40000000000000";
                    string Size_h_after = (778.4 * Size_chu).ToString();
                    Sizes.Attribute("height").Value = Size_h_after;
                    try
                    {
                        Sizes.Add(new XAttribute("isSetByUser", "true"));
                    }
                    catch (Exception)
                    {
                        break;
                    }
                    finally
                    {
                        Sizes.Attribute("isSetByUser").Value = "true";
                    }
                    
                    //int OEs_count = Outlines.Descendants(ns + "OE").Count();
                    //MessageBox.Show(OEs_count.ToString());
                    //if (OEs_count > 1)
                    //{
                    //    foreach (XElement OEs in from node in Outlines.Descendants(ns + "OE") select node)
                    //    {
                    //        string OEs_Styles = OEs.Attribute("style").Value;
                    //        if (String.IsNullOrEmpty(OEs_Styles))
                    //        {
                    //            OEs.Add(new XAttribute("style", "font-family:宋体;font-size:14.0pt;color:black"));
                    //        }
                    //        else
                    //        {
                    //            OEs.Attribute("style").Value= "font-family:宋体;font-size:14.0pt;color:black";
                    //        }
                    //    }
                    //    foreach (XElement Ts_1 in from node in Outlines.Descendants(ns + "T") select node)
                    //    {
                    //        string Ts_1_Styles = Ts_1.Attribute("style").Value;
                    //        if (String.IsNullOrEmpty(Ts_1_Styles))
                    //        {
                    //            Ts_1.Add(new XAttribute("style", "font-family:宋体;font-size:14.0pt;color:black"));
                    //        }
                    //        else
                    //        {
                    //            Ts_1.Attribute("style").Value = "font-family:宋体;font-size:14.0pt;color:black";
                    //        }
                    //    }
                    //    foreach (XElement Numbers in from node in Outlines.Descendants(ns + "Number") select node)
                    //    {
                    //        string Numbers_fontColors = Numbers.Attribute("fontColor").Value;
                    //        if (String.IsNullOrEmpty(Numbers_fontColors))
                    //        {
                    //            Numbers.Add(new XAttribute("fontColor", "#000000"));
                    //        }
                    //        else
                    //        {
                    //            Numbers.Attribute("fontColor").Value= "#000000";
                    //        }
                    //        string Numbers_fontSizes = Numbers.Attribute("fontSize").Value;
                    //        if (String.IsNullOrEmpty(Numbers_fontSizes))
                    //        {
                    //            Numbers.Add(new XAttribute("fontSize", "14.0"));
                    //        }
                    //        else
                    //        {
                    //            Numbers.Attribute("fontSize").Value= "14.0";
                    //        }
                    //        string Numbers_fonts = Numbers.Attribute("font").Value;
                    //        if (String.IsNullOrEmpty(Numbers_fontSizes))
                    //        {
                    //            Numbers.Add(new XAttribute("font", "宋体"));
                    //        }
                    //        else
                    //        {
                    //            Numbers.Attribute("font").Value = "宋体";
                    //        }
                    //    }
                    //}
                    //else
                    //{
                    //    XElement OEs_1 = Outlines.Descendants(ns + "OE").FirstOrDefault();
                    //    string Styles_1 = OEs_1.Attribute("style").Value;
                    //    if (String.IsNullOrEmpty(Styles_1))
                    //    {
                    //        OEs_1.Add(new XAttribute("style", "font-family:宋体;font-size:14.0pt;color:black"));
                    //    }
                    //    else
                    //    {
                    //        OEs_1.Attribute("style").Value = "font-family:宋体;font-size:14.0pt;color:black";
                    //    }
                    //}
                }
            }
            onenoteApp.UpdatePageContent(doc.ToString(), System.DateTime.MinValue);
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
