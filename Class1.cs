using Extensibility;
using Microsoft.Office.Core;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Threading;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using OneNote = Microsoft.Office.Interop.OneNote;
using TextCopy;
using System.Text;
using System.Text.RegularExpressions;
using HtmlAgilityPack;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Microsoft.Office.Tools.Ribbon;



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

        

        public static  void tittle_first()
        {
            string xmlPageTitle;
            OneNote.Application onenoteApp = new OneNote.Application();
            onenoteApp.GetPageContent(onenoteApp.Windows.CurrentWindow.CurrentPageId, out xmlPageTitle, OneNote.PageInfo.piAll);
            var xmlDoc = XDocument.Parse(xmlPageTitle);
            XNamespace ns = xmlDoc.Root.Name.Namespace;
            XElement OutLine_title = xmlDoc.Descendants(ns + "Title").FirstOrDefault();
            string outLine_titles_one = OutLine_title.Descendants(ns + "OE").FirstOrDefault().Attribute("objectID").Value;
            onenoteApp.NavigateTo(onenoteApp.Windows.CurrentWindow.CurrentPageId, outLine_titles_one, false);
        }
        public static void update_tittle_all()
        {
            tittle_first();
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
                    continue;
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
                            continue;
                        }
                        else if (outLine_titles_all.Value.Contains(outLine_tag) == true)
                        {
                            continue;
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
        public static void Del_tags(IRibbonControl control, string p_name)
        {
            tittle_first();
            OneNote.Application onenoteApp = new OneNote.Application();
            string xml;
            var pageid = onenoteApp.Windows.CurrentWindow.CurrentPageId;
            onenoteApp.GetPageContent(pageid, out xml, OneNote.PageInfo.piAll);
            var doc = XDocument.Parse(xml);
            XNamespace ns = doc.Root.Name.Namespace;
            try
            {
                XElement Indexs = doc.Descendants(ns + "TagDef").Where(xe => xe.Attribute("name").Value == p_name).FirstOrDefault();
            }
            catch (Exception)
            {

            }
            finally
            {
                XElement Indexs = doc.Descendants(ns + "TagDef").Where(xe => xe.Attribute("name").Value == p_name).FirstOrDefault();
                string p_index = Indexs.Attribute("index").Value;
                doc.Descendants(ns + "Tag").Where(xe => xe.Attribute("index") != null && xe.Attribute("index").Value == p_index).Remove();
                doc.Descendants(ns + "TagDef").Where(xe => xe.Attribute("name") != null && xe.Attribute("name").Value == p_name).Remove();
                onenoteApp.UpdatePageContent(doc.ToString(), System.DateTime.MinValue);
                update_tittle_all();
            }


        }
        public void Del_all_tags(IRibbonControl control)
        {
            tittle_first();
            OneNote.Application onenoteApp = new OneNote.Application();
            string xml;
            var pageid = onenoteApp.Windows.CurrentWindow.CurrentPageId;
            onenoteApp.GetPageContent(pageid, out xml, OneNote.PageInfo.piAll);
            var doc = XDocument.Parse(xml);
            XNamespace ns = doc.Root.Name.Namespace;
            try
            {
                XElement Indexs = doc.Descendants(ns + "TagDef").Where(xe => xe.Attribute("name") != null).FirstOrDefault();
            }
            catch (Exception)
            {

            }
            finally
            {
                foreach (XElement TagDefs in doc.Descendants(ns + "TagDef").ToList())
                {
                    TagDefs.Remove();
                }
                foreach (XElement Tags in doc.Descendants(ns + "Tag").ToList())
                {
                    Tags.Remove();
                }
                onenoteApp.UpdatePageContent(doc.ToString(), System.DateTime.MinValue);
                update_tittle_all();
            }


        }
        public static void Set_tags(IRibbonControl control, string p_type, string p_name)
        {
            tittle_first();
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
        public void Playlist_add_kaizhanzhong(IRibbonControl control)
        {
            Set_tags(control, "10", "【开展中】");
        }
        public void Playlist_del_kaizhanzhong(IRibbonControl control)
        {
            Del_tags(control, "【开展中】");
        }
        public void playlist_add_add(IRibbonControl control)
        {
            Set_tags(control, "11", "【未开展】");
        }
        public void playlist_del_add(IRibbonControl control)
        {
            Del_tags(control, "【未开展】");
        }
        public void Playlist_add_weiqueren(IRibbonControl control)
        {
            Set_tags(control, "12", "【未确认】");
        }
        public void Playlist_del_weiqueren(IRibbonControl control)
        {
            Del_tags(control, "【未确认】");
        }
        public void Playlist_add_zuofei(IRibbonControl control)
        {
            Set_tags(control, "13", "【作废】");
        }
        public void Playlist_del_zuofei(IRibbonControl control)
        {
            Del_tags(control, "【作废】");
        }
        public void Playlist_add_daisheji(IRibbonControl control)
        {
            Set_tags(control, "14", "【待设计】");
        }
        public void Playlist_del_daisheji(IRibbonControl control)
        {
            Del_tags(control, "【待设计】");
        }
        public void Playlist_add_weizhuan(IRibbonControl control)
        {
            Set_tags(control, "15", "【未转】");
        }
        public void Playlist_del_weizhuan(IRibbonControl control)
        {
            Del_tags(control, "【未转】");
        }
        public void Playlist_add_hebing(IRibbonControl control)
        {
            Set_tags(control, "16", "【合并】");
        }
        public void Playlist_del_hebing(IRibbonControl control)
        {
            Del_tags(control, "【合并】");
        }
        public void Playlist_add_yizhuan(IRibbonControl control)
        {
            Set_tags(control, "17", "【已转】");
        }
        public void Playlist_del_yizhuan(IRibbonControl control)
        {
            Del_tags(control, "【已转】");
        }
        public void Playlist_add_zanbukaizhan(IRibbonControl control)
        {
            Set_tags(control, "18", "【暂不开展】");
        }
        public void Playlist_del_zanbukaizhan(IRibbonControl control)
        {
            Del_tags(control, "【暂不开展】");
        }

        public void Playlist_add_yizhuanxubuchong(IRibbonControl control)
        {
            Set_tags(control, "19", "【已转需补充】");
        }
        public void Playlist_del_yizhuanxubuchong(IRibbonControl control)
        {
            Del_tags(control, "【已转需补充】");
        }
        public void Playlist_add_yiwancheng(IRibbonControl control)
        {
            Set_tags(control, "20", "【已完成】");
        }
        public void Playlist_del_yiwancheng(IRibbonControl control)
        {
            Del_tags(control, "【已完成】");
        }
        public void Playlist_add_p_work_n(IRibbonControl control)
        {
            Set_tags(control, "21", "项目工作（内部）");
        }
        public void Playlist_del_p_work_n(IRibbonControl control)
        {
            Del_tags(control, "项目工作（内部）");
        }

        public void Playlist_add_p_work_w(IRibbonControl control)
        {
            Set_tags(control, "22", "项目工作（外部）");
        }
        public void Playlist_del_p_work_w(IRibbonControl control)
        {
            Del_tags(control, "项目工作（外部）");
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
                if (String.IsNullOrEmpty(OutLine_data) && (OutLine_count == 1))
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
            Del_none(control);
            OneNote.Application onenoteApp = new OneNote.Application();
            string xml;
            var pageid = onenoteApp.Windows.CurrentWindow.CurrentPageId;
            onenoteApp.GetPageContent(pageid, out xml, OneNote.PageInfo.piAll);
            var doc = XDocument.Parse(xml);
            XNamespace ns = doc.Root.Name.Namespace;
            foreach (XElement Outlines in from node in doc.Descendants(ns + "Outline") select node)
            {
                //MessageBox.Show(Outlines.ToString());
                string OutLine_data = Outlines.Descendants(ns + "T").FirstOrDefault().Value.ToString();
                if (String.IsNullOrEmpty(OutLine_data))
                {
                    continue;
                }
                else 
                {
                    //MessageBox.Show(OutLine_data);
                    XElement Page_Meta = doc.Descendants(ns + "Meta").Where(x => x.Attribute("name").Value == "TaggingKit.PageTags").FirstOrDefault();
                    String Mate_content;
                    if (Page_Meta != null)
                    {
                        Mate_content = Page_Meta.Attribute("content").Value;
                    }
                    else
                    {
                        Mate_content = null;
                    }
                    if ((OutLine_data.Replace(" ", "") == Mate_content) && Mate_content != null)
                    {
                        continue;

                    }
                    else
                    {
                        XElement OutLine_Meta_Isnull = Outlines.Descendants(ns + "Meta").FirstOrDefault();
                        if (OutLine_Meta_Isnull != null)
                        {
                            foreach (XElement OutLine_Metas in from node1 in Outlines.Descendants(ns + "Meta") select node1)
                            {
                                string OutLine_Meta = OutLine_Metas.Attribute("name").Value;
                                if (OutLine_Meta != "omTaggingBank")
                                {

                                    XElement Positions = Outlines.Descendants(ns + "Position").FirstOrDefault();
                                    Positions.Attribute("x").Value = "36.00000000000000";
                                    Positions.Attribute("y").Value = "86.4000015258789";
                                    XElement Sizes = Outlines.Descendants(ns + "Size").FirstOrDefault();
                                    string Size_w = Sizes.Attribute("width").Value;
                                    string Size_h = Sizes.Attribute("height").Value;
                                    double Size_w_int = double.Parse(Size_w);
                                    double Size_h_int = double.Parse(Size_h);
                                    //769.8897094726562
                                    double Size_old = Size_w_int * Size_h_int;
                                    if (Size_old < 347432.4203514568)
                                    {
                                        Sizes.Attribute("width").Value = "451.2755737304687";
                                        Sizes.Attribute("height").Value = "769.8897094726562";
                                    }
                                    else
                                    {
                                        double Size_chu = Size_h_int / Size_w_int;
                                        Sizes.Attribute("width").Value = "451.2755737304687";
                                        string Size_h_after = (451.2755737304687 * Size_chu).ToString();
                                        Sizes.Attribute("height").Value = Size_h_after;
                                    }
                                    try
                                    {
                                        Sizes.Add(new XAttribute("isSetByUser", "true"));
                                    }
                                    catch (Exception)
                                    {

                                    }
                                    finally
                                    {
                                        Sizes.Attribute("isSetByUser").Value = "true";
                                    }
                                    onenoteApp.UpdatePageContent(doc.ToString(), System.DateTime.MinValue);
                                }


                            }
                        }
                        else
                        {
                            XElement Positions = Outlines.Descendants(ns + "Position").FirstOrDefault();
                            Positions.Attribute("x").Value = "36.00000000000000";
                            Positions.Attribute("y").Value = "86.4000015258789";
                            XElement Sizes = Outlines.Descendants(ns + "Size").FirstOrDefault();
                            string Size_w = Sizes.Attribute("width").Value;
                            string Size_h = Sizes.Attribute("height").Value;
                            double Size_w_int = double.Parse(Size_w);
                            double Size_h_int = double.Parse(Size_h);
                            //769.8897094726562
                            double Size_old = Size_w_int * Size_h_int;
                            if (Size_old < 347432.4203514568)
                            {
                                Sizes.Attribute("width").Value = "451.2755737304687";
                                Sizes.Attribute("height").Value = "769.8897094726562";
                            }
                            else
                            {
                                double Size_chu = Size_h_int / Size_w_int;
                                Sizes.Attribute("width").Value = "451.2755737304687";
                                string Size_h_after = (451.2755737304687 * Size_chu).ToString();
                                Sizes.Attribute("height").Value = Size_h_after;
                            }
                            try
                            {
                                Sizes.Add(new XAttribute("isSetByUser", "true"));
                            }
                            catch (Exception)
                            {

                            }
                            finally
                            {
                                Sizes.Attribute("isSetByUser").Value = "true";
                            }
                            onenoteApp.UpdatePageContent(doc.ToString(), System.DateTime.MinValue);
                        }
                        


                    }
                
                   
                    
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
        public void riji_create_page(IRibbonControl control)
        {
            // 获取当前日期
            var date = DateTime.Now.ToString("yyyy-MM-dd");
            var date_year= DateTime.Now.ToString("yyyy");
            var date_mouth= DateTime.Now.ToString("MMMM", new CultureInfo("zh-CN"));

            var application = new OneNote.Application();
            String onenote_file;
            application.GetSpecialLocation((OneNote.SpecialLocation)2, out onenote_file);
            // Get info from OneNote 
            string xml;
            application.GetHierarchy(null, OneNote.HierarchyScope.hsSections, out xml);
            XDocument doc = XDocument.Parse(xml);
            XNamespace ns = doc.Root.Name.Namespace;

            
            // Assuming you have a notebook called "Test" 
            XElement notebook = doc.Root.Elements(ns + "Notebook").Where(x => x.Attribute("name").Value == "My Journal").FirstOrDefault();
            if (notebook == null)
            {
                String strID_1;
                String notebook_string;
                //MessageBox.Show(onenote_file + "\\My Project Journal\\" + date_year + "\\" + date_mouth + ".one");
                application.OpenHierarchy(onenote_file + "\\My Journal\\",
                System.String.Empty, out strID_1, OneNote.CreateFileType.cftNotebook);
                application.GetHierarchy(strID_1, OneNote.HierarchyScope.hsNotebooks, out notebook_string);
                notebook = XElement.Parse(notebook_string);
            }

            

            // If there is a section, just use the first one we encounter 
            XElement section_year = notebook.Elements(ns + "SectionGroup").Where(x => x.Attribute("name").Value == date_year).FirstOrDefault();
            if (section_year == null)
            {
                String strID_2;
                String section_year_string;
                string strID_1 = notebook.Attribute("ID").Value;
                application.OpenHierarchy(date_year, strID_1, out strID_2, OneNote.CreateFileType.cftFolder);
                application.GetHierarchy(strID_2, OneNote.HierarchyScope.hsSections, out section_year_string);
                section_year = XElement.Parse(section_year_string);
            }


            XElement section_mouth = section_year.Elements(ns + "Section").Where(x => x.Attribute("name").Value == date_mouth).FirstOrDefault();
            if (section_mouth == null)
            {
                String strID_3;
                String section_mouth_string;
                string strID_2 = section_year.Attribute("ID").Value;
                application.OpenHierarchy(date_mouth + ".one", strID_2, out strID_3, OneNote.CreateFileType.cftSection);
                application.GetHierarchy(strID_3, OneNote.HierarchyScope.hsSections, out section_mouth_string);
                section_mouth = XElement.Parse(section_mouth_string);

            }


            // Create a page 
            XElement section_day = section_mouth.Elements(ns + "Section").Where(x => x.Attribute("name").Value == date).FirstOrDefault();
            if (section_day == null)
            {
                string newPageID;
                application.CreateNewPage(section_mouth.Attribute("ID").Value, out newPageID);

                //MessageBox.Show(newPageID);
                // Create the page element using the ID of the new page OneNote just created 
                XElement newPage = new XElement(ns + "Page");
                newPage.SetAttributeValue("ID", newPageID);


                // Add a title just for grins 
                newPage.Add(new XElement(ns + "Title",
                    new XElement(ns + "OE",
                     new XElement(ns + "T",
                      new XCData(date)))));
                // Add an outline and text content 
                newPage.Add(new XElement(ns + "Outline",
                    new XElement(ns + "OEChildren",
                     new XElement(ns + "OE",
                      new XElement(ns + "T",
                       new XCData(""))))));
                // 创建 OneNote 页
                //MessageBox.Show(newPage.ToString());
                application.UpdatePageContent(newPage.ToString());
            }


        }

       


        public void create_my_ribao_kong(IRibbonControl control)
        {
            // 获取当前日期
            var date = DateTime.Now.ToString("yyyy-MM-dd");
            string[] weekdays = { "星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六" };
            string wkCN = weekdays[Convert.ToInt32(DateTime.Now.DayOfWeek)];
            var date_day = date + " " + wkCN;
            var date_year = DateTime.Now.ToString("yyyy");
            var date_mouth = DateTime.Now.ToString("MMMM", new CultureInfo("zh-CN"));

            var application = new OneNote.Application();
            String onenote_file;
            application.GetSpecialLocation((OneNote.SpecialLocation)2, out onenote_file);
            // Get info from OneNote 
            string xml;
            application.GetHierarchy(null, OneNote.HierarchyScope.hsSections, out xml);
            XDocument doc = XDocument.Parse(xml);
            XNamespace ns = doc.Root.Name.Namespace;
            // Assuming you have a notebook called "Test" 
            XElement notebook = doc.Root.Elements(ns + "Notebook").Where(x => x.Attribute("name").Value == "My Work Log").FirstOrDefault();
            if (notebook == null)
            {
                String strID_1;
                String notebook_string;
                //MessageBox.Show(onenote_file + "\\My Project Journal\\" + date_year + "\\" + date_mouth + ".one");
                application.OpenHierarchy(onenote_file + "\\My Work Log\\",
                System.String.Empty, out strID_1, OneNote.CreateFileType.cftNotebook);
                application.GetHierarchy(strID_1, OneNote.HierarchyScope.hsNotebooks, out notebook_string);
                notebook = XElement.Parse(notebook_string);
            }



            // If there is a section, just use the first one we encounter 
            XElement section_year = notebook.Elements(ns + "SectionGroup").Where(x => x.Attribute("name").Value == date_year).FirstOrDefault();
            if (section_year == null)
            {
                String strID_2;
                String section_year_string;
                string strID_1 = notebook.Attribute("ID").Value;
                application.OpenHierarchy(date_year, strID_1, out strID_2, OneNote.CreateFileType.cftFolder);
                application.GetHierarchy(strID_2, OneNote.HierarchyScope.hsSections, out section_year_string);
                section_year = XElement.Parse(section_year_string);
            }


            XElement section_mouth = section_year.Elements(ns + "Section").Where(x => x.Attribute("name").Value == date_mouth).FirstOrDefault();
            if (section_mouth == null)
            {
                String strID_3;
                String section_mouth_string;
                string strID_2 = section_year.Attribute("ID").Value;
                application.OpenHierarchy(date_mouth + ".one", strID_2, out strID_3, OneNote.CreateFileType.cftSection);
                application.GetHierarchy(strID_3, OneNote.HierarchyScope.hsSections, out section_mouth_string);
                section_mouth = XElement.Parse(section_mouth_string);

            }


            // Create a page 
            XElement section_day = section_mouth.Elements(ns + "Section").Where(x => x.Attribute("name").Value == date_day).FirstOrDefault();
            if (section_day == null)
            {
                string newPageID;
                application.CreateNewPage(section_mouth.Attribute("ID").Value, out newPageID);

                //MessageBox.Show(newPageID);
                // Create the page element using the ID of the new page OneNote just created 
                XElement newPage = new XElement(ns + "Page");
                newPage.SetAttributeValue("ID", newPageID);



                // Add a title just for grins 
                newPage.Add(new XElement(ns + "Title",
                    new XElement(ns + "OE",
                     new XElement(ns + "T",
                      new XCData(date_day)))));
                // Add an outline and text content 
                newPage.Add(new XElement(ns + "Outline",
                    new XElement(ns + "OEChildren",
                     new XElement(ns + "OE",
                      new XElement(ns + "T",
                       new XCData(""))))));
                // 创建 OneNote 页
                //MessageBox.Show(newPage.ToString());
                application.UpdatePageContent(newPage.ToString());
            }



        }
        public void create_my_ribao_quan(IRibbonControl control)
        {
            // 获取当前日期
            var date = DateTime.Now.ToString("yyyy-MM-dd");
            string[] weekdays = { "星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六" };
            string wkCN = weekdays[Convert.ToInt32(DateTime.Now.DayOfWeek)];
            var date_day = date + " " + wkCN;
            var date_year = DateTime.Now.ToString("yyyy");
            var date_mouth = DateTime.Now.ToString("MMMM", new CultureInfo("zh-CN"));

            var application = new OneNote.Application();
            String onenote_file;
            application.GetSpecialLocation((OneNote.SpecialLocation)2, out onenote_file);
            // Get info from OneNote 
            string xml;
            application.GetHierarchy(null, OneNote.HierarchyScope.hsSections, out xml);
            XDocument doc = XDocument.Parse(xml);
            XNamespace ns = doc.Root.Name.Namespace;
            // Assuming you have a notebook called "Test" 
            XElement notebook = doc.Root.Elements(ns + "Notebook").Where(x => x.Attribute("name").Value == "My Work Log").FirstOrDefault();
            if (notebook == null)
            {
                String strID_1;
                String notebook_string;
                //MessageBox.Show(onenote_file + "\\My Project Journal\\" + date_year + "\\" + date_mouth + ".one");
                application.OpenHierarchy(onenote_file + "\\My Work Log\\",
                System.String.Empty, out strID_1, OneNote.CreateFileType.cftNotebook);
                application.GetHierarchy(strID_1, OneNote.HierarchyScope.hsNotebooks, out notebook_string);
                notebook = XElement.Parse(notebook_string);
            }



            // If there is a section, just use the first one we encounter 
            XElement section_year = notebook.Elements(ns + "SectionGroup").Where(x => x.Attribute("name").Value == date_year).FirstOrDefault();
            if (section_year == null)
            {
                String strID_2;
                String section_year_string;
                string strID_1 = notebook.Attribute("ID").Value;
                application.OpenHierarchy(date_year, strID_1, out strID_2, OneNote.CreateFileType.cftFolder);
                application.GetHierarchy(strID_2, OneNote.HierarchyScope.hsSections, out section_year_string);
                section_year = XElement.Parse(section_year_string);
            }


            XElement section_mouth = section_year.Elements(ns + "Section").Where(x => x.Attribute("name").Value == date_mouth).FirstOrDefault();
            if (section_mouth == null)
            {
                String strID_3;
                String section_mouth_string;
                string strID_2 = section_year.Attribute("ID").Value;
                application.OpenHierarchy(date_mouth + ".one", strID_2, out strID_3, OneNote.CreateFileType.cftSection);
                application.GetHierarchy(strID_3, OneNote.HierarchyScope.hsSections, out section_mouth_string);
                section_mouth = XElement.Parse(section_mouth_string);

            }


            // Create a page 
            XElement section_day = section_mouth.Elements(ns + "Section").Where(x => x.Attribute("name").Value == date_day).FirstOrDefault();
            if (section_day == null)
            {
                string newPageID;
                application.CreateNewPage(section_mouth.Attribute("ID").Value, out newPageID);

                //MessageBox.Show(newPageID);
                // Create the page element using the ID of the new page OneNote just created 
                XElement newPage = new XElement(ns + "Page");
                newPage.SetAttributeValue("ID", newPageID);

                // Add a title just for grins 
                newPage.Add(new XElement(ns + "Title",
                        new XElement(ns + "OE",
                         new XElement(ns + "T",
                          new XCData(date_day)))));
                // Add an outline and text content 
                /*newPage.Add(new XElement(ns + "Outline",
                    new XElement(ns + "OEChildren",
                     new XElement(ns + "OE",
                      new XElement(ns + "T",
                       new XCData(""))))));*/
                newPage.Add(new XElement(ns + "Outline",
                       new XElement(ns + "OEChildren",
                        new XElement(ns + "OE",
                        new XElement(ns + "Table",
                        new XAttribute("bordersVisible", "true"),
                        new XAttribute("hasHeaderRow", "true")
                        ,
                        new XElement(ns + "Columns",
                        new XElement(ns + "Column",
                        new XAttribute("index", "0"),
                        new XAttribute("width", "37.11000061035156")

                        ),
                         new XElement(ns + "Column",
                        new XAttribute("index", "1"),
                        new XAttribute("width", "402.9299926757812"),
                        new XAttribute("isLocked", "true")

                        )
                         ),
                        new XElement(ns + "Row",
                        new XElement(ns + "Cell",
                        new XElement(ns + "OEChildren",
                        new XElement(ns + "OE",
                         new XElement(ns + "T",
                          new XCData("日报")
                          )
                         )
                        )
                        ),
                        new XElement(ns + "Cell",
                        new XElement(ns + "OEChildren",
                        new XElement(ns + "OE",
                        new XElement(ns + "Table",
                        new XAttribute("bordersVisible", "true"),
                        new XAttribute("hasHeaderRow", "true")
                        ,
                        new XElement(ns + "Columns",
                        new XElement(ns + "Column",
                        new XAttribute("index", "0"),
                        new XAttribute("width", "48.20996856689453")

                        ),
                         new XElement(ns + "Column",
                        new XAttribute("index", "1"),
                        new XAttribute("width", "340.1100463867187"),
                        new XAttribute("isLocked", "true")

                        )
                         ),
                        new XElement(ns + "Row",
                        new XElement(ns + "Cell",
                        new XElement(ns + "OEChildren",
                        new XElement(ns + "OE",
                         new XElement(ns + "T",
                          new XCData("工作计划")
                         )
                        )
                        )
                        ),
                        new XElement(ns + "Cell",
                        new XElement(ns + "OEChildren",
                        new XElement(ns + "OE",
                         new XElement(ns + "T",
                          new XCData("")
                         )
                        )
                        )
                        )
                        ),
                        new XElement(ns + "Row",
                        new XElement(ns + "Cell",
                        new XElement(ns + "OEChildren",
                        new XElement(ns + "OE",
                         new XElement(ns + "T",
                          new XCData("工作结果")
                         )
                        )
                        )
                        ),
                        new XElement(ns + "Cell",
                        new XElement(ns + "OEChildren",
                        new XElement(ns + "OE",
                         new XElement(ns + "T",
                          new XCData("")
                         )
                        )
                        )
                        )
                        ),
                        new XElement(ns + "Row",
                        new XElement(ns + "Cell",
                        new XElement(ns + "OEChildren",
                        new XElement(ns + "OE",
                         new XElement(ns + "T",
                          new XCData("佐证资料")
                         )
                        )
                        )
                        ),
                        new XElement(ns + "Cell",
                        new XElement(ns + "OEChildren",
                        new XElement(ns + "OE",
                         new XElement(ns + "T",
                          new XCData("")
                         )
                        )
                        )
                        )
                        )
                        )
                        )
                        )

                        )
                        )
                         )
                        )
                        )
                       )
                    );
                // 创建 OneNote 页
                //MessageBox.Show(newPage.ToString());
                application.UpdatePageContent(newPage.ToString());
            }

        }


        public void create_my_ribao_dan(IRibbonControl control)
        {
            // 获取当前日期
            var date = DateTime.Now.ToString("yyyy-MM-dd");
            string[] weekdays = { "星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六" };
            string wkCN = weekdays[Convert.ToInt32(DateTime.Now.DayOfWeek)];
            var date_day = date + " " + wkCN;
            var date_year = DateTime.Now.ToString("yyyy");
            var date_mouth = DateTime.Now.ToString("MMMM", new CultureInfo("zh-CN"));

            var application = new OneNote.Application();
            String onenote_file;
            application.GetSpecialLocation((OneNote.SpecialLocation)2, out onenote_file);
            // Get info from OneNote 
            string xml;
            application.GetHierarchy(null, OneNote.HierarchyScope.hsSections, out xml);
            XDocument doc = XDocument.Parse(xml);
            XNamespace ns = doc.Root.Name.Namespace;
            // Assuming you have a notebook called "Test" 
            XElement notebook = doc.Root.Elements(ns + "Notebook").Where(x => x.Attribute("name").Value == "My Work Log").FirstOrDefault();
            if (notebook == null)
            {
                String strID_1;
                String notebook_string;
                //MessageBox.Show(onenote_file + "\\My Project Journal\\" + date_year + "\\" + date_mouth + ".one");
                application.OpenHierarchy(onenote_file + "\\My Work Log\\",
                System.String.Empty, out strID_1, OneNote.CreateFileType.cftNotebook);
                application.GetHierarchy(strID_1, OneNote.HierarchyScope.hsNotebooks, out notebook_string);
                notebook = XElement.Parse(notebook_string);
            }



            // If there is a section, just use the first one we encounter 
            XElement section_year = notebook.Elements(ns + "SectionGroup").Where(x => x.Attribute("name").Value == date_year).FirstOrDefault();
            if (section_year == null)
            {
                String strID_2;
                String section_year_string;
                string strID_1 = notebook.Attribute("ID").Value;
                application.OpenHierarchy(date_year, strID_1, out strID_2, OneNote.CreateFileType.cftFolder);
                application.GetHierarchy(strID_2, OneNote.HierarchyScope.hsSections, out section_year_string);
                section_year = XElement.Parse(section_year_string);
            }


            XElement section_mouth = section_year.Elements(ns + "Section").Where(x => x.Attribute("name").Value == date_mouth).FirstOrDefault();
            if (section_mouth == null)
            {
                String strID_3;
                String section_mouth_string;
                string strID_2 = section_year.Attribute("ID").Value;
                application.OpenHierarchy(date_mouth + ".one", strID_2, out strID_3, OneNote.CreateFileType.cftSection);
                application.GetHierarchy(strID_3, OneNote.HierarchyScope.hsSections, out section_mouth_string);
                section_mouth = XElement.Parse(section_mouth_string);

            }


            // Create a page 
            XElement section_day = section_mouth.Elements(ns + "Section").Where(x => x.Attribute("name").Value == date_day).FirstOrDefault();
            if (section_day == null)
            {
                string newPageID;
                application.CreateNewPage(section_mouth.Attribute("ID").Value, out newPageID);

                //MessageBox.Show(newPageID);
                // Create the page element using the ID of the new page OneNote just created 
                XElement newPage = new XElement(ns + "Page");
                newPage.SetAttributeValue("ID", newPageID);

                // Add a title just for grins 
                newPage.Add(new XElement(ns + "Title",
                        new XElement(ns + "OE",
                         new XElement(ns + "T",
                          new XCData(date_day)))));
                // Add an outline and text content 
                /*newPage.Add(new XElement(ns + "Outline",
                    new XElement(ns + "OEChildren",
                     new XElement(ns + "OE",
                      new XElement(ns + "T",
                       new XCData(""))))));*/
                newPage.Add(new XElement(ns + "Outline",
                       new XElement(ns + "OEChildren",
                        new XElement(ns + "OE",
                        new XElement(ns + "Table",
                        new XAttribute("bordersVisible", "true"),
                        new XAttribute("hasHeaderRow", "true")
                        ,
                        new XElement(ns + "Columns",
                        new XElement(ns + "Column",
                        new XAttribute("index", "0"),
                        new XAttribute("width", "37.11000061035156")

                        ),
                         new XElement(ns + "Column",
                        new XAttribute("index", "1"),
                        new XAttribute("width", "402.9299926757812"),
                        new XAttribute("isLocked", "true")

                        )
                         ),

                        new XElement(ns + "Row",
                        new XElement(ns + "Cell",
                        new XElement(ns + "OEChildren",
                        new XElement(ns + "OE",
                         new XElement(ns + "T",
                          new XCData("日报")
                          )
                         )
                        )
                        ),
                        new XElement(ns + "Cell",
                        new XElement(ns + "OEChildren",
                        new XElement(ns + "OE",
                        new XElement(ns + "Table",
                        new XAttribute("bordersVisible", "true"),
                        new XAttribute("hasHeaderRow", "true")
                        ,
                       new XElement(ns + "Columns",
                        new XElement(ns + "Column",
                        new XAttribute("index", "0"),
                        new XAttribute("width", "48.20996856689453")

                        ),
                         new XElement(ns + "Column",
                        new XAttribute("index", "1"),
                        new XAttribute("width", "340.1100463867187"),
                        new XAttribute("isLocked", "true")

                        )
                         ),

                        new XElement(ns + "Row",
                        new XElement(ns + "Cell",
                        new XElement(ns + "OEChildren",
                        new XElement(ns + "OE",
                         new XElement(ns + "T",
                          new XCData("工作结果")
                         )
                        )
                        )
                        ),
                        new XElement(ns + "Cell",
                        new XElement(ns + "OEChildren",
                        new XElement(ns + "OE",
                         new XElement(ns + "T",
                          new XCData("")
                         )
                        )
                        )
                        )
                        )
                       
                        
                        
                        
                        )
                        )
                        )
                        )

                        )
                        )
                         )
                        )
                        )
                       
                    
                    );
                // 创建 OneNote 页
                //MessageBox.Show(newPage.ToString());
                application.UpdatePageContent(newPage.ToString());
            }

        }


        public static void create_ribao(IRibbonControl control, string Page_Notebook, string Page_SectionGroup, string Page_Section, string Page_Tittle,string Page_type=null)
        {
            

            var application = new OneNote.Application();
            String onenote_file;
            application.GetSpecialLocation((OneNote.SpecialLocation)2, out onenote_file);
            // Get info from OneNote 
            string xml;
            application.GetHierarchy(null, OneNote.HierarchyScope.hsSections, out xml);
            XDocument doc = XDocument.Parse(xml);
            XNamespace ns = doc.Root.Name.Namespace;
            // Assuming you have a notebook called "Test" 
            XElement notebook = doc.Root.Elements(ns + "Notebook").Where(x => x.Attribute("name").Value == Page_Notebook).FirstOrDefault();
            if (notebook == null)
            {
                String strID_1;
                String notebook_string;
                application.OpenHierarchy(onenote_file + "\\"+ Page_Notebook + "\\",
                System.String.Empty, out strID_1, OneNote.CreateFileType.cftNotebook);
                application.GetHierarchy(strID_1, OneNote.HierarchyScope.hsNotebooks, out notebook_string);
                notebook = XElement.Parse(notebook_string);
            }

            
            // If there is a section, just use the first one we encounter 
            XElement section_year = notebook.Elements(ns + "SectionGroup").Where(x => x.Attribute("name").Value == Page_SectionGroup).FirstOrDefault();
            if (section_year == null)
            {
                String strID_2;
                String section_year_string;
                string strID_1 = notebook.Attribute("ID").Value;
                application.OpenHierarchy(Page_SectionGroup, strID_1, out strID_2, OneNote.CreateFileType.cftFolder);
                application.GetHierarchy(strID_2, OneNote.HierarchyScope.hsSections, out section_year_string);
                section_year = XElement.Parse(section_year_string);
            }


            XElement section_mouth = section_year.Elements(ns + "Section").Where(x => x.Attribute("name").Value == Page_Section).FirstOrDefault();
            if (section_mouth == null)
            {
                String strID_3;
                String section_mouth_string;
                string strID_2 = section_year.Attribute("ID").Value;
                application.OpenHierarchy(Page_Section + ".one", strID_2, out strID_3, OneNote.CreateFileType.cftSection);
                application.GetHierarchy(strID_3, OneNote.HierarchyScope.hsSections, out section_mouth_string);
                section_mouth = XElement.Parse(section_mouth_string);

            }


            // Create a page 
            XElement section_day = section_mouth.Elements(ns + "Section").Where(x => x.Attribute("name").Value == Page_Tittle).FirstOrDefault();
            if (section_day == null)
            {
                string newPageID;
                application.CreateNewPage(section_mouth.Attribute("ID").Value, out newPageID);

                //MessageBox.Show(newPageID);
                // Create the page element using the ID of the new page OneNote just created 
                XElement newPage = new XElement(ns + "Page");
                newPage.SetAttributeValue("ID", newPageID);

                // Add a title just for grins 
                newPage.Add(new XElement(ns + "Title",
                        new XElement(ns + "OE",
                         new XElement(ns + "T",
                          new XCData(Page_Tittle)))));
                // Add an outline and text content 
                /*newPage.Add(new XElement(ns + "Outline",
                    new XElement(ns + "OEChildren",
                     new XElement(ns + "OE",
                      new XElement(ns + "T",
                       new XCData(""))))));*/
                if (Page_type==null || Page_type== "riji")
                {
                    newPage.Add(new XElement(ns + "Outline",
                    new XElement(ns + "OEChildren",
                     new XElement(ns + "OE",
                      new XElement(ns + "T",
                       new XCData(""))))));
                }
                else if (Page_type == "ribao_quan")
                {
                    newPage.Add(new XElement(ns + "Outline",
                       new XElement(ns + "OEChildren",
                        new XElement(ns + "OE",
                        new XElement(ns + "Table",
                        new XAttribute("bordersVisible", "true"),
                        new XAttribute("hasHeaderRow", "true")
                        ,
                        new XElement(ns + "Columns",
                        new XElement(ns + "Column",
                        new XAttribute("index", "0"),
                        new XAttribute("width", "37.11000061035156")

                        ),
                         new XElement(ns + "Column",
                        new XAttribute("index", "1"),
                        new XAttribute("width", "402.9299926757812"),
                        new XAttribute("isLocked", "true")

                        )
                         ),
                        new XElement(ns + "Row",
                        new XElement(ns + "Cell",
                        new XElement(ns + "OEChildren",
                        new XElement(ns + "OE",
                         new XElement(ns + "T",
                          new XCData("日报")
                          )
                         )
                        )
                        ),
                        new XElement(ns + "Cell",
                        new XElement(ns + "OEChildren",
                        new XElement(ns + "OE",
                        new XElement(ns + "Table",
                        new XAttribute("bordersVisible", "true"),
                        new XAttribute("hasHeaderRow", "true")
                        ,
                        new XElement(ns + "Columns",
                        new XElement(ns + "Column",
                        new XAttribute("index", "0"),
                        new XAttribute("width", "48.20996856689453")

                        ),
                         new XElement(ns + "Column",
                        new XAttribute("index", "1"),
                        new XAttribute("width", "340.1100463867187"),
                        new XAttribute("isLocked", "true")

                        )
                         ),
                        new XElement(ns + "Row",
                        new XElement(ns + "Cell",
                        new XElement(ns + "OEChildren",
                        new XElement(ns + "OE",
                         new XElement(ns + "T",
                          new XCData("工作计划")
                         )
                        )
                        )
                        ),
                        new XElement(ns + "Cell",
                        new XElement(ns + "OEChildren",
                        new XElement(ns + "OE",
                         new XElement(ns + "T",
                          new XCData("")
                         )
                        )
                        )
                        )
                        ),
                        new XElement(ns + "Row",
                        new XElement(ns + "Cell",
                        new XElement(ns + "OEChildren",
                        new XElement(ns + "OE",
                         new XElement(ns + "T",
                          new XCData("工作结果")
                         )
                        )
                        )
                        ),
                        new XElement(ns + "Cell",
                        new XElement(ns + "OEChildren",
                        new XElement(ns + "OE",
                         new XElement(ns + "T",
                          new XCData("")
                         )
                        )
                        )
                        )
                        ),
                        new XElement(ns + "Row",
                        new XElement(ns + "Cell",
                        new XElement(ns + "OEChildren",
                        new XElement(ns + "OE",
                         new XElement(ns + "T",
                          new XCData("佐证资料")
                         )
                        )
                        )
                        ),
                        new XElement(ns + "Cell",
                        new XElement(ns + "OEChildren",
                        new XElement(ns + "OE",
                         new XElement(ns + "T",
                          new XCData("")
                         )
                        )
                        )
                        )
                        )
                        )
                        )
                        )

                        )
                        )
                         )
                        )
                        )
                       )
                    );
                }
                else if (Page_type == "ribao_dan")
                {
                    newPage.Add(new XElement(ns + "Outline",
                       new XElement(ns + "OEChildren",
                        new XElement(ns + "OE",
                        new XElement(ns + "Table",
                        new XAttribute("bordersVisible", "true"),
                        new XAttribute("hasHeaderRow", "true")
                        ,
                        new XElement(ns + "Columns",
                        new XElement(ns + "Column",
                        new XAttribute("index", "0"),
                        new XAttribute("width", "37.11000061035156")

                        ),
                         new XElement(ns + "Column",
                        new XAttribute("index", "1"),
                        new XAttribute("width", "402.9299926757812"),
                        new XAttribute("isLocked", "true")

                        )
                         ),

                        new XElement(ns + "Row",
                        new XElement(ns + "Cell",
                        new XElement(ns + "OEChildren",
                        new XElement(ns + "OE",
                         new XElement(ns + "T",
                          new XCData("日报")
                          )
                         )
                        )
                        ),
                        new XElement(ns + "Cell",
                        new XElement(ns + "OEChildren",
                        new XElement(ns + "OE",
                        new XElement(ns + "Table",
                        new XAttribute("bordersVisible", "true"),
                        new XAttribute("hasHeaderRow", "true")
                        ,
                       new XElement(ns + "Columns",
                        new XElement(ns + "Column",
                        new XAttribute("index", "0"),
                        new XAttribute("width", "48.20996856689453")

                        ),
                         new XElement(ns + "Column",
                        new XAttribute("index", "1"),
                        new XAttribute("width", "340.1100463867187"),
                        new XAttribute("isLocked", "true")

                        )
                         ),

                        new XElement(ns + "Row",
                        new XElement(ns + "Cell",
                        new XElement(ns + "OEChildren",
                        new XElement(ns + "OE",
                         new XElement(ns + "T",
                          new XCData("工作结果")
                         )
                        )
                        )
                        ),
                        new XElement(ns + "Cell",
                        new XElement(ns + "OEChildren",
                        new XElement(ns + "OE",
                         new XElement(ns + "T",
                          new XCData("")
                         )
                        )
                        )
                        )
                        )




                        )
                        )
                        )
                        )

                        )
                        )
                         )
                        )
                        )


                    );
                }
                else
                {
                    newPage.Add(new XElement(ns + "Outline",
                    new XElement(ns + "OEChildren",
                     new XElement(ns + "OE",
                      new XElement(ns + "T",
                       new XCData(""))))));
                }
                
                // 创建 OneNote 页
                //MessageBox.Show(newPage.ToString());
                application.UpdatePageContent(newPage.ToString());
            }

        }

        public void create_my_ribao_quan_all(IRibbonControl control)
        {
           
            Application.EnableVisualStyles();
            //Application.SetCompatibleTextRenderingDefault(false);
            DateForm box1 = new DateForm();
            Application.Run(box1);
            box1.button1.PerformClick();
            DateTime start_date = box1.Start_Date.Value;
            DateTime stop_date = box1.Last_Date.Value;
            List<DateTime> listDay = new List<DateTime>();
            DateTime dtDay = new DateTime();
            //循环比较，取出日期；
            for (dtDay = start_date; dtDay.CompareTo(stop_date) <= 0; dtDay = dtDay.AddDays(1))
            {
                listDay.Add(dtDay);
            }
            foreach (DateTime today in listDay)
            {
                var date = today.ToString("yyyy-MM-dd");
                string[] weekdays = { "星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六" };
                string wkCN = weekdays[Convert.ToInt32(today.DayOfWeek)];
                var date_day = date + " " + wkCN;
                var date_year = today.ToString("yyyy");
                var date_mouth = today.ToString("MMMM", new CultureInfo("zh-CN"));
                create_ribao(control, "My Work Log", date_year, date_mouth, date_day, "ribao_quan");
            }



            
        }
        public void create_my_ribao_quan_leater(IRibbonControl control)
        {
            // 获取当前日期
            var date = DateTime.Now.AddDays(1).ToString("yyyy-MM-dd");
            string[] weekdays = { "星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六" };
            string wkCN = weekdays[Convert.ToInt32(DateTime.Now.AddDays(1).DayOfWeek)];
            var date_day = date + " " + wkCN;
            var date_year = DateTime.Now.AddDays(1).ToString("yyyy");
            var date_mouth = DateTime.Now.AddDays(1).ToString("MMMM", new CultureInfo("zh-CN"));           
            create_ribao(control, "My Work Log", date_year, date_mouth, date_day, "ribao_quan");
        }
        public void page_a4(IRibbonControl control)
        {

            OneNote.Application onenoteApp = new OneNote.Application();
            string xml;
            var pageid = onenoteApp.Windows.CurrentWindow.CurrentPageId;
            onenoteApp.GetPageContent(pageid, out xml, OneNote.PageInfo.piAll);
            var doc = XDocument.Parse(xml);
            XNamespace ns = doc.Root.Name.Namespace;
           
            XElement oldPageSize = doc.Descendants(ns + "PageSize").FirstOrDefault();
            //MessageBox.Show(TagDefs.ToString());
            oldPageSize.RemoveNodes();
            XElement newPageSize_Ori = new XElement(ns + "Orientation",
                                            new XAttribute("landscape", "false")
                                            );
            XElement newPageSize_Dim = new XElement(ns + "Dimensions",
                                            new XAttribute("width", "595.2755737304687"),
                                            new XAttribute("height", "841.8897094726562")
                                            );
            XElement newPageSize_Mar = new XElement(ns + "Margins",
                                            new XAttribute("top", "36.0"),
                                            new XAttribute("bottom", "36.0"),
                                            new XAttribute("left", "72.0"),
                                            new XAttribute("right", "72.0")
                                            );
            //MessageBox.Show(newTagDefs.ToString());
            oldPageSize.Add(newPageSize_Ori);
            oldPageSize.Add(newPageSize_Dim);
            oldPageSize.Add(newPageSize_Mar);


            //MessageBox.Show(doc.ToString());
            onenoteApp.UpdatePageContent(doc.ToString());
           
        }
        
        public class notebooks_list
        {
            public string notebooks_list_data;
        }
        public static void pruject_name(IRibbonControl control,out List<notebooks_list> list)
        {

            OneNote.Application onenoteApp = new OneNote.Application();
            string notebook_xml;
            onenoteApp.GetHierarchy("", OneNote.HierarchyScope.hsNotebooks, out notebook_xml);
            
            XDocument doc = XDocument.Parse(notebook_xml);
            XNamespace ns = doc.Root.Name.Namespace;
            // Assuming you have a notebook called "Test" 
            var notebooks_list = new List<notebooks_list>();
            foreach (XElement notebook in from node1 in doc.Root.Elements(ns + "Notebook") select node1)
            {
                string notebooks_list_datas = notebook.Attribute("name").Value;
                
                notebooks_list.Add(new notebooks_list() { notebooks_list_data = notebooks_list_datas });
            }
            
            list = notebooks_list;
            
        }

       
        public void create_xuqiu_page(IRibbonControl control)
        {
           
            OneNote.Application onenoteApp = new OneNote.Application();
            string outLine_SectionId = onenoteApp.Windows.CurrentWindow.CurrentSectionId;
            string outLine_PageId;
            onenoteApp.CreateNewPage(outLine_SectionId,out outLine_PageId);
            string xml;
            onenoteApp.GetPageContent(outLine_PageId, out xml, OneNote.PageInfo.piAll);
            var doc = XDocument.Parse(xml);
            XNamespace ns = doc.Root.Name.Namespace;

            // 获取当前日期
            var date = DateTime.Now.ToString("yyyy-MM-dd");
            var date_year = DateTime.Now.ToString("yyyy");
            var date_mouth = DateTime.Now.ToString("MMMM", new CultureInfo("zh-CN"));


            // Get info from OneNote 



            XElement newPage = new XElement(ns + "Page");
            newPage.SetAttributeValue("ID", outLine_PageId);

            // Add a title just for grins 
            newPage.Add(new XElement(ns + "Title",
                new XElement(ns + "OE",
                 new XElement(ns + "T",
                  new XCData("**｜"+date)))));
            // Add an outline and text content 
            newPage.Add(new XElement(ns + "Outline",
                new XElement(ns + "OEChildren",
                 new XElement(ns + "OE",
                  new XElement(ns + "T",
                   new XCData(""))))));
            // 创建 OneNote 页
            
            onenoteApp.UpdatePageContent(newPage.ToString());
            //MessageBox.Show(newPage.ToString());



        }

        public  int GetItemCount(IRibbonControl control)
        {
            List<notebooks_list> ItemLabels;
            pruject_name(control, out ItemLabels);
            return ItemLabels.Count;

        }
        public string GetItemLabel(IRibbonControl control, int index)
        {
            List<notebooks_list> ItemLabels ;
            pruject_name(control, out ItemLabels);
            return ItemLabels[index].notebooks_list_data.ToString();
            
        }
        public string GetItemID(IRibbonControl control, int index)
        {
            return "heading" + index.ToString();
        }

        public void GetonAction(IRibbonControl control)
        {

        }

        public static void get_chun(out String page_copy)
        {
            OneNote.Application onenoteApp = new OneNote.Application();
            string xml;
            var pageid = onenoteApp.Windows.CurrentWindow.CurrentPageId;
            onenoteApp.GetPageContent(pageid, out xml, OneNote.PageInfo.piAll);
            XDocument doc = XDocument.Parse(xml);
            XNamespace ns = doc.Root.Name.Namespace;
            StringBuilder sb = new StringBuilder();
            foreach (XElement Outlines in doc.Descendants(ns + "Outline").ToList())
            {
                String Outlines_selected = null;
                Outlines_selected = Outlines.Attribute("selected").Value;
                if (Outlines_selected != null)
                {
                    foreach (XElement Outlines_OEChilds in Outlines.Descendants(ns + "OEChildren").ToList())
                    {
                        String OEChilds_selected = null;
                        try
                        {
                            OEChilds_selected = Outlines_OEChilds.Attribute("selected").Value;
                        }
                        catch (Exception)
                        {

                        }

                        if (OEChilds_selected != null)
                        {
                            //MessageBox.Show(Outlines_OEChilds.ToString());
                            foreach (XElement Outlines_OE in Outlines_OEChilds.Descendants(ns + "OE").ToList())
                            {


                                String OE_selected = null;
                                try
                                {
                                    OE_selected = Outlines_OE.Attribute("selected").Value;
                                }
                                catch (Exception)
                                {

                                }

                                if (OE_selected != null)
                                {
                                    //MessageBox.Show(Outlines_OE.ToString());
                                    foreach (XElement Outlines_T in Outlines_OE.Descendants(ns + "T").ToList())
                                    {
                                        String T_selected = null;
                                        try
                                        {
                                            T_selected = Outlines_T.Attribute("selected").Value;
                                        }
                                        catch (Exception)
                                        {

                                        }

                                        if (T_selected != null)
                                        {
                                            MessageBox.Show(Outlines_T.ToString());
                                            sb.AppendLine(Outlines_T.Value);
                                        }
                                        else
                                        {
                                            continue;
                                        }
                                    }
                                }
                                else
                                {

                                    continue;
                                }
                            }
                        }
                        else
                        {
                            continue;
                        }
                    }
                }
                else
                {
                    continue;
                }


            }
            page_copy= sb.ToString();
            /*MessageBox.Show(sb.ToString());
            TextCopy.ClipboardService.SetText(sb.ToString());*/
        }

        public static void Replace_cpan(String cpan_in,out String cpan_out)
        {



            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            doc.LoadHtml(cpan_in);
            StringBuilder sb_1 = new StringBuilder();
            var spanNodes = doc.DocumentNode.SelectNodes("//span");
            if (spanNodes != null)
            {
                foreach (var spanNode in spanNodes)
                {
                    string content = spanNode.InnerHtml;
                    sb_1.Append(content);
                }
            }
            cpan_out= sb_1.ToString();

        }

        public void copy_chun(IRibbonControl control)
        {
            OneNote.Application onenoteApp = new OneNote.Application();
            string xml;
            var pageid = onenoteApp.Windows.CurrentWindow.CurrentPageId;
            onenoteApp.GetPageContent(pageid, out xml, OneNote.PageInfo.piAll);
            XDocument doc = XDocument.Parse(xml);
            XNamespace ns = doc.Root.Name.Namespace;
            StringBuilder sb = new StringBuilder();
            String page_copy;

            foreach (XElement Outlines_T in doc.Descendants(ns + "T").ToList())
            {
                String T_selected = null;
                try
                {
                    T_selected = Outlines_T.Attribute("selected").Value;
                }
                catch (Exception)
                {

                }

                if (T_selected == "all")
                {
                    //MessageBox.Show(Outlines_T.ToString());
                    //sb.AppendLine(Outlines_T.Value);
                    if(Outlines_T.Value.Contains("span") ==true)
                    {
                        Replace_cpan(Outlines_T.Value, out page_copy);
                        sb.AppendLine(page_copy);
                    }
                    else
                    {
                        sb.AppendLine(Outlines_T.Value);
                    }
                    
                }
                else
                {
                    continue;
                }
            }
            //MessageBox.Show(sb.ToString());
            TextCopy.ClipboardService.SetText(sb.ToString());
            MessageBox.Show("复制成功！");
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
