using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using Aspose.Pdf.Devices;
using System.Collections;
using Aspose.Pdf.Text;
using System.Globalization;
using System.Speech.Synthesis;
using WaveLib;
using Yeti.MMedia.Mp3;
using System.Text.RegularExpressions;


namespace ePubFix
{
    
    public partial class MForm : ComponentFactory.Krypton.Toolkit.KryptonForm
    {
        
        string inPDFpath = "";
        string curErrLog = "";
        string outCopyPath = "";
        int inPDFPageCount = 0;
        BackgroundWorker _bw;
        delegate void update_probar(string text,int max,int cur);
        bool audioComplete = false;

        string spanTxt = "";
        string smilTxt = "";
        string last_time = "";
        int spanID = 1;
        string curPageHtml = "";
        string curaudioFile = "";
        string curFullTxt = "";
        string htLineTxt = "";
        int spanPoint = 0;
        string voiceName = "";
        ArrayList cpList = new ArrayList();
        ArrayList attList = new ArrayList();

        public MForm()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                this.Text += " " + Application.ProductVersion;
               
                gCls.update_path_var();
                cmb_aupage.SelectedIndex = 0;
                txt_aupagetxt.Visible = false;
                load_voice_list();
                chk_audio.Checked = false;
                cmb_voicelist.Enabled = false;
            }
            catch (Exception erd) {
               gCls.show_error(erd.Message.ToString());
                return;
            }
        }

        public void load_voice_list() {

            cmb_voicelist.Items.Clear();
            using (SpeechSynthesizer synthesizer = new SpeechSynthesizer())
            {
                // Output information about all of the installed voices that
                // support the en-US locacale. 
              
                foreach (InstalledVoice voice in  synthesizer.GetInstalledVoices(new CultureInfo("en-US")))
                {
                    VoiceInfo info = voice.VoiceInfo;
                    cmb_voicelist.Items.Add(info.Name);
                }
                
            }
            if (cmb_voicelist.Items.Count > 0) {
                cmb_voicelist.SelectedIndex = 0;
            }            
        
        }
        

        public void probar_update(string Etext, int max, int cur) {
            try
            {
                if (probar.InvokeRequired)
                {
                    update_probar up = new update_probar(probar_update);
                    this.Invoke(up, new object[] { Etext, max, cur });
                }
                else
                {
                    if (Etext != "")
                    {
                        lbstatus.Text = Etext;
                    }
                    probar.Maximum = max;
                    probar.Value = cur;
                }

            }
            catch { }
        
        }

        private void kryptonGroupBox1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void kryptonGroupBox1_Panel_Paint(object sender, PaintEventArgs e)
        {

        }

        private void kryptonPanel1_Paint(object sender, PaintEventArgs e)
        {
            
        }

        private void kryptonButton5_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void kryptonButton4_Click(object sender, EventArgs e)
        {
            cpList.Clear();
            submit_convert();
        }



        public void submit_convert() {

            try
            {
                if (txt_inpdf.Text == "") {
                    gCls.show_error("Select PDF file");
                    return;
                }
               

                if (!File.Exists(txt_inpdf.Text)) {
                    gCls.show_error("PDF File not found");
                    return;
                }

                if (check_kf.Checked) {
                    if (txt_booktitle.Text == "" || txt_bookauthor.Text == "") {
                        gCls.show_error("Enter book title and author information");
                        return;
                    }

                }

                if (txt_epubcover.Text != "")
                {
                    if (!File.Exists(txt_epubcover.Text)) {
                        gCls.show_error("ePub cover file not found");
                        return;
                    }

                    if (Path.GetExtension(txt_epubcover.Text).ToLower() != ".jpg") {
                        gCls.show_error("ePub cover should be jpg file format.");
                        return;
                    }
                }

                #region check audio custom page
                if (chk_audio.Checked) {
                    if (cmb_aupage.Text == "Custom Page")
                    {
                        if (txt_aupagetxt.Text == "")
                        {
                            gCls.show_error("Enter custom page number");
                            return;
                        }
                        cpList.Clear();
                        try
                        {

                            string txtPage = txt_aupagetxt.Text;
                            txtPage = txtPage.Replace(" ", "");

                            #region validation
                            char[] charX = txtPage.ToCharArray();
                            foreach (char c in charX)
                            {
                                if (c.ToString() != "," && c.ToString() != "-" && char.IsNumber(c) == false)
                                {
                                    gCls.show_error("Invalid custom page value");
                                    return;
                                }
                            }
                            #endregion

                            #region add into first list
                            ArrayList firstList = new ArrayList();
                            if (txtPage.IndexOf(",") != -1)
                            {
                                string[] tflist = txtPage.Split(',');
                                foreach (string tf in tflist)
                                {
                                    firstList.Add(tf);
                                }
                            }
                            else
                            {
                                firstList.Add(txtPage);
                            }

                            #endregion

                            #region get pages from array
                            for (int i = 0; i < firstList.Count; i++)
                            {
                                string uEnt = firstList[i].ToString();
                                if (uEnt.IndexOf("-") != -1)
                                {
                                    string[] uTxt = uEnt.Split('-');
                                    int s1 = int.Parse(uTxt[0]);
                                    int s2 = int.Parse(uTxt[1]);
                                    if (s1 >= s2)
                                    {
                                        gCls.show_error("Invalid custom page value");
                                        return;
                                    }
                                    for (int v = s1; v <= s2; v++)
                                    {

                                        cpList.Add(v.ToString());
                                    }
                                }
                                else
                                {

                                    cpList.Add(uEnt);
                                }
                            }
                            #endregion

                        }
                        catch (Exception erd)
                        {
                            gCls.show_error("Invalid custom page : " + erd.Message.ToString());
                            return;
                        }

                    }
                    else { //select all page count
                        cpList.Clear();
                        aspose_license_update();
                        Aspose.Pdf.Document pdfDocx = new Aspose.Pdf.Document(txt_inpdf.Text);
                      
                        for (int p = 1; p <= pdfDocx.Pages.Count; p++) {
                            cpList.Add(p);
                        }
                    
                    }                     
                
                }


                #endregion

                #region check crop value
                

                #endregion

                outCopyPath = Path.GetDirectoryName(txt_inpdf.Text);

                    

                string b_filename = Path.GetFileNameWithoutExtension(txt_inpdf.Text);
                string b_title = txt_booktitle.Text;
                string b_author = txt_bookauthor.Text;
                FileInfo fInfo = new FileInfo(txt_inpdf.Text);
                string b_filesize = fInfo.Length.ToString();
                int b_resolution = Convert.ToInt32(txt_dpi.Value);
                string b_pagetype = "normal";
                string userfolder = Application.StartupPath +  "\\in_pdf";
                File.Copy(txt_inpdf.Text, Application.StartupPath + "\\in_pdf\\" + b_filename + ".pdf", true);

                if (File.Exists(Application.StartupPath + "\\in_pdf\\cover.jpg"))
                {
                    File.Delete(Application.StartupPath + "\\in_pdf\\cover.jpg");
                }

                if (txt_epubcover.Text != "") {
                    
                    File.Copy(txt_epubcover.Text, Application.StartupPath + "\\in_pdf\\cover.jpg", true);
                }
                bool b_kf = check_kf.Checked;
                Updf live_pdf_info = new Updf(b_filename, b_title, b_author, b_filesize, b_resolution.ToString(), userfolder, b_pagetype,b_kf);
                curErrLog = "";
                inPDFPageCount = 0;
                 _bw = new BackgroundWorker
                {
                    WorkerReportsProgress = true,
                    WorkerSupportsCancellation = true
                };
                _bw.DoWork += bw_DoWork;
                probar.Visible = true;
                _bw.ProgressChanged += bw_ProgressChanged;
                _bw.RunWorkerCompleted += bw_RunWorkerCompleted;

                //voice name
                voiceName = cmb_voicelist.Text;
                _bw.RunWorkerAsync(live_pdf_info);
                
               
            }
            catch (Exception erd) {
                MessageBox.Show(erd.Message.ToString());
                probar.Visible = false;
                return;
            }
        
        }
        private void bw_DoWork(object sender, DoWorkEventArgs e)
        {
            Updf live_pdf_info = (Updf)e.Argument;           
            convert_start(live_pdf_info);     
           
        }
        private void bw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            probar.Visible = false;
            lbstatus.Text = "";
            if (curErrLog != "")
            {
                gCls.show_error(curErrLog);
                return;
            }
            else
            {
                string epubFile = inPDFpath.Substring(0, inPDFpath.Length - 3) + "epub";
                string mobiFile = inPDFpath.Substring(0, inPDFpath.Length - 3) + "mobi";
                string eFilename = Path.GetFileName(epubFile);
                string mFilename = Path.GetFileName(mobiFile);
                if (File.Exists(epubFile)) {
                    File.Copy(epubFile, outCopyPath + "\\" + eFilename, true);
                }
                
                if (File.Exists(mobiFile))
                {
                    File.Copy(mobiFile, outCopyPath + "\\" + mFilename, true);
                }
                
                inPDFPageCount = 0;
                gCls.show_message("ePub converted successfully.");
            }
        }
        private void bw_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            
        }


        public bool check_audio_custom_page(int pNum) {

            try
            {
                bool aFound = false;
                for (int i = 0; i < cpList.Count; i++) {
                   int pV = int.Parse(cpList[i].ToString());
                   if (pV == pNum) { aFound = true; }
                }
                return aFound;
            }
            catch {
                return false;
            }
        
        }

        public void SaveStreamToFile(string fileFullPath, Stream stream)
        {
            try
            {
                if (stream.Length == 0) return;
                using (FileStream fileStream = System.IO.File.Create(fileFullPath, (int)stream.Length))
                {
                    byte[] bytesInStream = new byte[stream.Length];
                    stream.Read(bytesInStream, 0, (int)bytesInStream.Length);
                    fileStream.Write(bytesInStream, 0, bytesInStream.Length);
                }
            }
            catch { }
        }


        public void convert_start(Updf cPDF)
        {

            curErrLog = "";

            try
            {
                aspose_license_update();
                string b_filename = cPDF.b_filename;
                inPDFpath = cPDF.usrfolder + "\\" + cPDF.b_filename + ".pdf";
                string b_outfilepath = cPDF.usrfolder;
                Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(inPDFpath);


                #region crop pdf
               if(chk_crop.Checked){
                   Double crop_left = Convert.ToDouble(txt_crop_left.Value);
                   Double crop_right = Convert.ToDouble(txt_crop_right.Value);
                   Double crop_top = Convert.ToDouble(txt_crop_top.Value);
                   Double crop_bottom = Convert.ToDouble(txt_crop_bottom.Value);
                   
                for (int c = 1; c < pdfDocument.Pages.Count + 1; c++)
                {
                   Double crop_width = pdfDocument.Pages[c].Rect.Width - (crop_left + crop_right);
                   Double crop_height = pdfDocument.Pages[c].Rect.Height - (crop_top + crop_bottom);
                   
                    Aspose.Pdf.Rectangle pageRect = new Aspose.Pdf.Rectangle(crop_left,crop_top,crop_width,crop_height);
                    pdfDocument.Pages[c].CropBox = pageRect;                
                }

                pdfDocument.Save(inPDFpath);
                pdfDocument = new Aspose.Pdf.Document(inPDFpath);  
               }

                #endregion

                string doc_title = "";
                string doc_author = "";
                if (cPDF.b_title != "") { doc_title = cPDF.b_title; }
                else { doc_title = pdfDocument.Info.Title; }
                if (cPDF.b_author != "") { doc_author = cPDF.b_author; }
                else { doc_author = pdfDocument.Info.Author; }
                int imgDPI = Convert.ToInt32(cPDF.b_resolution);

                #region cover image extract from page1
                //cover page epub
                if (!File.Exists(cPDF.usrfolder + "\\cover.jpg"))
                {
                    using (FileStream imageStream = new FileStream(cPDF.usrfolder + "\\cover.jpg", FileMode.Create))
                    {
                        Resolution resolution = new Resolution(imgDPI);
                        JpegDevice jpegDevice = new JpegDevice(resolution, 100);
                        jpegDevice.Process(pdfDocument.Pages[1], imageStream);
                        imageStream.Close();
                    }
                }
                //cover page kf
                if (cPDF.b_kf)
                {
                    if (!File.Exists(cPDF.usrfolder + "\\cover_kf.jpg"))
                    {
                        using (FileStream imageStream = new FileStream(cPDF.usrfolder + "\\cover_kf.jpg", FileMode.Create))
                        {
                            Resolution resolution = new Resolution(96);
                            JpegDevice jpegDevice = new JpegDevice(resolution, 100);
                            jpegDevice.Process(pdfDocument.Pages[1], imageStream);
                            imageStream.Close();
                        }
                    }
                }
                #endregion

                int fontID = 1;
                ArrayList fntCollection = new ArrayList();
                ArrayList pageCollection = new ArrayList();

                probar_update("Text extract...", pdfDocument.Pages.Count, 0);


                #region each page


                for (int i = 1; i < pdfDocument.Pages.Count + 1; i++)
                {
                    probar_update("", pdfDocument.Pages.Count, i);


                    ArrayList txCollection = new ArrayList();
                    TextFragmentAbsorber textFragmentAbsorber;
                    if (chk_wordposition.Checked)
                    {
                        textFragmentAbsorber = new TextFragmentAbsorber(@"(?<=^|\s)" + "(.+?)"  + @"(?=\s|$)");
                        Aspose.Pdf.Text.TextOptions.TextSearchOptions textSearchOptions = new Aspose.Pdf.Text.TextOptions.TextSearchOptions(true);                        
                        textFragmentAbsorber.TextSearchOptions = textSearchOptions;
                    }
                    else {
                        textFragmentAbsorber = new TextFragmentAbsorber();
                       
                    }

                    try
                    {
                        pdfDocument.Pages[i].Accept(textFragmentAbsorber);
                    }
                    catch {
                        textFragmentAbsorber = new TextFragmentAbsorber();
                        pdfDocument.Pages[i].Accept(textFragmentAbsorber);
                    }

                    TextFragmentCollection textFragmentCollection = textFragmentAbsorber.TextFragments;
                    foreach (TextFragment textFragment in textFragmentCollection)
                    {

                        string f_text = "";
                        string f_left = "";
                        string f_top = "";
                        string font_name = "";
                        string font_size = "";
                        string font_color = "";

                        #region word position
                        //if (chk_wordposition.Checked) // word coordinate
                        //{
                        //    foreach (Aspose.Pdf.Text.TextSegment sg in textFragment.Segments) {
                        //        f_text = sg.Text;
                        //        f_left = (sg.BaselinePosition.XIndent - pdfDocument.Pages[i].CropBox.LLX).ToString();
                        //        double f_top1 = (pdfDocument.Pages[i].CropBox.Height - sg.BaselinePosition.YIndent) - textFragment.Rectangle.Height;
                        //        f_top = (f_top1 + pdfDocument.Pages[i].CropBox.LLY).ToString();
                        //        font_name = sg.TextState.Font.FontName;
                        //        font_size = sg.TextState.FontSize.ToString();
                        //        font_color = gCls.HexConverter(sg.TextState.ForegroundColor);

                        //        #region add into array
                        //        bool fntFound = false;
                        //        string fntID = "";
                        //        for (int f = 0; f < fntCollection.Count; f++)
                        //        {
                        //            fInfo fl = (fInfo)fntCollection[f];
                        //            if (fl.f_color == font_color && fl.f_fontfamily == font_name && fl.f_fontsize == font_size)
                        //            {
                        //                fntID = fl.f_fontid;
                        //                fntFound = true;
                        //            }
                        //        }

                        //        if (fntFound == false)
                        //        {
                        //            fontID++;
                        //            fInfo nfl = new fInfo(font_size, font_name, font_color, "fnt" + fontID.ToString());
                        //            fntCollection.Add(nfl);
                        //            fntID = "fnt" + fontID.ToString();
                        //        }

                        //        tInfo txm = new tInfo(f_left, f_top, f_text, fntID);
                        //        txCollection.Add(txm);
                               

                        //        #endregion

                        //    } // end each word loop

                        //    try
                        //    {
                        //        textFragment.Text = "";
                        //    }
                        //    catch { }

                        //}
                        #endregion
                       
                            f_text = textFragment.Text;
                            f_left = (textFragment.BaselinePosition.XIndent - pdfDocument.Pages[i].CropBox.LLX).ToString();
                            double f_top1 = (pdfDocument.Pages[i].CropBox.Height - textFragment.BaselinePosition.YIndent) - textFragment.Rectangle.Height;
                            f_top = (f_top1 + pdfDocument.Pages[i].CropBox.LLY).ToString();
                            font_name = textFragment.TextState.Font.FontName;
                            font_size = textFragment.TextState.FontSize.ToString();
                            font_color = gCls.HexConverter(textFragment.TextState.ForegroundColor);

                            #region add into array
                            bool fntFound = false;
                            string fntID = "";
                            for (int f = 0; f < fntCollection.Count; f++)
                            {
                                fInfo fl = (fInfo)fntCollection[f];
                                if (fl.f_color == font_color && fl.f_fontfamily == font_name && fl.f_fontsize == font_size)
                                {
                                    fntID = fl.f_fontid;
                                    fntFound = true;
                                }
                            }

                            if (fntFound == false)
                            {
                                fontID++;
                                fInfo nfl = new fInfo(font_size, font_name, font_color, "fnt" + fontID.ToString());
                                fntCollection.Add(nfl);
                                fntID = "fnt" + fontID.ToString();
                            }

                            tInfo txm = new tInfo(f_left, f_top, f_text, fntID);
                            txCollection.Add(txm);

                            if (chk_wordposition.Checked == false)
                            {
                                try
                                {

                                    textFragment.Text = "";
                                }
                                catch { }
                            }

                            #endregion





                    }

                    #region delete text for word co-ordinate
                    if (chk_wordposition.Checked)
                    {
                        TextFragmentAbsorber nTextFrag = new TextFragmentAbsorber();                        
                        pdfDocument.Pages[i].Accept(nTextFrag);
                        TextFragmentCollection nTextFcol = nTextFrag.TextFragments;
                        foreach (TextFragment tfrt in nTextFcol)
                        {
                            try
                            {
                                tfrt.Text = "";
                            }
                            catch { }

                        }
                    }
                    #endregion

                    pageCollection.Add(txCollection);


                }//end page
                #endregion


                #region create directory
                //create directory
                string PDFfName = cPDF.b_filename;
                string wrkDir = cPDF.usrfolder + "\\" + PDFfName;
                string kfwrkDir = cPDF.usrfolder + "\\" + PDFfName + "_KF";
                string kfimgDir = "";
                string kfstyDir = "";
                string kfhtmlDir = "";
                #region create directory for KF
                if (cPDF.b_kf)
                {
                    if (!Directory.Exists(kfwrkDir))
                    {
                        Directory.CreateDirectory(kfwrkDir);
                    }

                    kfimgDir = kfwrkDir + "\\image";
                    kfstyDir = kfwrkDir + "\\css";
                    kfhtmlDir = kfwrkDir + "\\html";
                    if (!Directory.Exists(kfimgDir))
                    {
                        Directory.CreateDirectory(kfimgDir);
                    }
                    if (!Directory.Exists(kfstyDir))
                    {
                        Directory.CreateDirectory(kfstyDir);
                    }
                    if (!Directory.Exists(kfhtmlDir))
                    {
                        Directory.CreateDirectory(kfhtmlDir);
                    }


                }
                #endregion

                if (!Directory.Exists(wrkDir))
                {
                    Directory.CreateDirectory(wrkDir);
                }
                else { Directory.Delete(wrkDir, true);
                Directory.CreateDirectory(wrkDir);
                }
                //===========

                //create epub directory
                if (!Directory.Exists(wrkDir + "\\META-INF"))
                {
                    Directory.CreateDirectory(wrkDir + "\\META-INF");
                }
                if (!Directory.Exists(wrkDir + "\\OEBPS"))
                {
                    Directory.CreateDirectory(wrkDir + "\\OEBPS");
                }
                string oebDir = wrkDir + "\\OEBPS";
                if (!Directory.Exists(oebDir + "\\html"))
                {
                    Directory.CreateDirectory(oebDir + "\\html");
                }
                if (!Directory.Exists(oebDir + "\\images"))
                {
                    Directory.CreateDirectory(oebDir + "\\images");
                }
                if (!Directory.Exists(oebDir + "\\styles"))
                {
                    Directory.CreateDirectory(oebDir + "\\styles");
                }
                string imgDir = oebDir + "\\images";
                string styDir = oebDir + "\\styles";
                string htmlDir = oebDir + "\\html";
                
                //audio directory
                string audioDir = oebDir + "\\audio";
                if (chk_audio.Checked && !Directory.Exists(audioDir)) {
                    Directory.CreateDirectory(audioDir);
                }

                File.Copy(Application.StartupPath + "\\template\\META-INF\\com.apple.ibooks.display-options.xml", wrkDir + "\\META-INF\\com.apple.ibooks.display-options.xml", true);
                File.Copy(Application.StartupPath + "\\template\\META-INF\\container.xml", wrkDir + "\\META-INF\\container.xml", true);
                File.Copy(Application.StartupPath + "\\template\\mimetype", wrkDir + "\\mimetype", true);

               
                


                #endregion


                //====
                #region toc.ncx
                string tocX = File.ReadAllText(Application.StartupPath + "\\template\\OEBPS\\toc.ncx");
                tocX = tocX.Replace("<docTitle><text></text></docTitle>", "<docTitle><text>" + doc_title + "</text></docTitle>");
                tocX = tocX.Replace("<docAuthor><text></text></docAuthor>", "<docAuthor><text>" + doc_author + "</text></docAuthor>");
                tocX = tocX.Replace("<navLabel><text></text></navLabel>", "<navLabel><text>" + doc_title + "</text></navLabel>");
                File.WriteAllText(oebDir + "\\toc.ncx", tocX);

                #region kf toc.ncx
                if (cPDF.b_kf)
                {
                    tocX = File.ReadAllText(Application.StartupPath + "\\template\\OEBPS\\kftoc.ncx");
                    tocX = tocX.Replace("<docTitle><text></text></docTitle>", "<docTitle><text>" + doc_title + "</text></docTitle>");
                    tocX = tocX.Replace("<docAuthor><text></text></docAuthor>", "<docAuthor><text>" + doc_author + "</text></docAuthor>");
                    tocX = tocX.Replace("</text></navLabel>", doc_title + "</text></navLabel>");
                    File.WriteAllText(kfwrkDir + "\\toc.ncx", tocX);
                }

                #endregion


                #endregion


                attList.Clear();

                #region content.opf

                string contX = "<manifest>\n";
                contX += "<item id=\"cover\" href=\"html/cover.html\" media-type=\"application/xhtml+xml\" />\n";
                contX += "<item id=\"ncx\" href=\"toc.ncx\" media-type=\"application/x-dtbncx+xml\" />\n";

                //contX += "<item id=\"toc\" properties=\"nav\" href=\"html/toc.html\" media-type=\"application/xhtml+xml\" />\n";
                contX += "<item id=\"pagecommon\" href=\"styles/page.css\" media-type=\"text/css\" />\n";
                for (int p = 1; p < pdfDocument.Pages.Count + 1; p++)
                {
                    #region check attachment 
                    if(pdfDocument.Pages[p].Annotations.Count > 0){
                     
                        for(int an = 1; an <= pdfDocument.Pages[p].Annotations.Count; an++) {
                            try{
                                Aspose.Pdf.InteractiveFeatures.Annotations.FileAttachmentAnnotation fAttach = (Aspose.Pdf.InteractiveFeatures.Annotations.FileAttachmentAnnotation)pdfDocument.Pages[p].Annotations[an];
                                string attFileName = fAttach.File.Name;
                                string attExt = Path.GetExtension(attFileName).ToLower();
                                //video 
                                string atfullPath = imgDir + "\\page" + p.ToString() + "-" + an.ToString() + attExt;
                                string attFName = "page" + p.ToString() + "-" + an.ToString() + attExt;
                                SaveStreamToFile(atfullPath, fAttach.File.Contents);
                                string meType = "";
                                if (attExt == ".mp4" || attExt == ".ogv")
                                {
                                    meType = "";
                                    if (attExt == ".mp4"){meType = "video/mpeg";}
                                    else { meType = "video/ogg"; }                                    
                                    contX += "<item id=\"att" + p.ToString() + "-" + an.ToString() + "\" href=\"images/page" + p.ToString() + "-" + an.ToString() + attExt + "\" media-type=\"" + meType + "\" />\n";
                                }
                                else if(attExt == ".mp3" || attExt == ".ogg"){
                                    meType = "";
                                    if (attExt == ".mp3") { meType = "audio/mpeg"; }
                                    else { meType = "audio/ogg"; }                                    
                                    contX += "<item id=\"att" + p.ToString() + "-" + an.ToString() + "\" href=\"images/page" + p.ToString() + "-" + an.ToString() + attExt + "\" media-type=\"" + meType + "\" />\n";
                                }

                                string atX = ((fAttach.Rect.LLX * imgDPI) / 72).ToString();
                                string atY = (((pdfDocument.Pages[p].Rect.Height  - fAttach.Rect.LLY) * imgDPI) / 72) .ToString();
                                
                                atY = (Double.Parse(atY) - fAttach.Rect.Height).ToString(); 

                                string atWidth = "";
                                string atHeight = "";
                                string atDesc = fAttach.Contents;
                                if (Regex.Match(atDesc, "width=(\\d+)").Success) {
                                    atWidth = Regex.Match(atDesc, "width=(\\d+)").Groups[1].Value;
                                }
                                if (Regex.Match(atDesc, "height=(\\d+)").Success)
                                {
                                    atHeight = Regex.Match(atDesc, "height=(\\d+)").Groups[1].Value;
                                }
                                if (atWidth == "") { atWidth = "350"; }
                                if (atHeight == "") { atHeight = "250"; }

                                string mtType = "";
                                if (meType.IndexOf("audio") != -1) { mtType = "audio"; } else { mtType = "video"; }

                                attList.Add(p.ToString() + "|" + attFName + "|" + atX + "|" + atY + "|" + atWidth + "|" + atHeight + "|" + mtType);

                            }catch{}
                        }

                        //remove attachement
                        int atnCount = pdfDocument.Pages[p].Annotations.Count;
                        for (int atn = 0; atn < atnCount; atn++) {
                            try
                            {
                                pdfDocument.Pages[p].Annotations.Delete(1);
                            }
                            catch { }
                        }
                        
                    }
                

                    #endregion

                    if (chk_audio.Checked && check_audio_custom_page(p))
                    {
                        contX += "<item id=\"page" + p.ToString() + "\" href=\"html/page" + p.ToString() + ".html\" media-type=\"application/xhtml+xml\" media-overlay=\"page" + p.ToString() + "smil\" />\n";
                        contX += "<item id=\"page" + p.ToString() + "audio\" href=\"audio/audio" + p.ToString() + ".mp3\" media-type=\"audio/mpeg\" />\n";
                        contX += "<item id=\"page" + p.ToString() + "smil\" href=\"html/page" + p.ToString() + ".smil\" media-type=\"application/smil+xml\" />\n";                        
                    } else {
                        contX += "<item id=\"page" + p.ToString() + "\" href=\"html/page" + p.ToString() + ".html\" media-type=\"application/xhtml+xml\" />\n";
                    }
                 
                    contX += "<item id=\"image" + p.ToString() + "\" href=\"images/page" + p.ToString() + ".jpg\" media-type=\"image/jpeg\" />\n";
                    contX += "<item id=\"page" + p.ToString() + "css\" href=\"styles/page" + p.ToString() + ".css\" media-type=\"text/css\" />\n";
                }
                contX += "<item id=\"cover-image\" href=\"images/cover.jpg\" media-type=\"image/jpeg\" />\n";

               

                contX += "</manifest>\n";
                contX += "<spine toc=\"ncx\">\n";
                contX += "<itemref idref=\"cover\" />\n";
                for (int px = 1; px < pdfDocument.Pages.Count + 1; px++)
                {
                    contX += "<itemref idref=\"page" + px.ToString() + "\" />\n";
                }
                contX += "</spine>\n";

                string rcont = File.ReadAllText(Application.StartupPath + "\\template\\OEBPS\\content.opf");
                rcont = rcont.Replace("<dc:title></dc:title>", "<dc:title>" + doc_title + "</dc:title>");
                rcont = rcont.Replace("<dc:creator></dc:creator>", "<dc:creator>" + doc_author + "</dc:creator>");
                rcont = rcont.Replace("<dc:publisher></dc:publisher>", "<dc:publisher>" + doc_author + "</dc:publisher>");
                rcont = rcont.Replace("<dc:rights></dc:rights>", "<dc:rights>" + doc_author + "</dc:rights>");
                string mdyDate = DateTime.Today.ToString("yyyy-MM-dd");
                rcont = rcont.Replace("<dc:date></dc:date>", "<dc:date>" + mdyDate + "</dc:date>");

                rcont = rcont.Replace("</metadata>", "</metadata>" + contX);
                File.WriteAllText(oebDir + "\\content.opf", rcont);

                int kfiWidth = 0;
                int kfiHeight = 0;
                gCls.get_image_size(cPDF.usrfolder + "\\cover_kf.jpg", ref kfiWidth, ref kfiHeight);

                #region create kf content.opf file
                if (cPDF.b_kf)
                {
                    contX = "<manifest>\n";
                    contX += "<item id=\"cover\" href=\"html/cover.html\" media-type=\"application/xhtml+xml\" />\n";
                    contX += "<item id=\"ncx\" href=\"toc.ncx\" media-type=\"application/x-dtbncx+xml\" />\n";
                    contX += "<item id=\"pagecommon\" href=\"css/page.css\" media-type=\"text/css\" />\n";
                    for (int p = 1; p < pdfDocument.Pages.Count + 1; p++)
                    {
                        contX += "<item id=\"page" + p.ToString() + "\" href=\"html/page" + p.ToString() + ".html\" media-type=\"application/xhtml+xml\" />\n";
                        contX += "<item id=\"image" + p.ToString() + "\" href=\"image/page" + p.ToString() + ".jpg\" media-type=\"image/jpeg\" />\n";
                        contX += "<item id=\"page" + p.ToString() + "css\" href=\"css/page" + p.ToString() + ".css\" media-type=\"text/css\" />\n";
                    }
                    contX += "<item id=\"my_cover_image\" href=\"image/cover.jpg\" media-type=\"image/jpeg\" />\n";

                    contX += "</manifest>\n";
                    contX += "<spine toc=\"ncx\">\n";
                    contX += "<itemref idref=\"cover\" />\n";
                    for (int px = 1; px < pdfDocument.Pages.Count + 1; px++)
                    {
                        contX += "<itemref idref=\"page" + px.ToString() + "\" />\n";
                    }
                    contX += "</spine>\n";

                    rcont = File.ReadAllText(Application.StartupPath + "\\template\\OEBPS\\kfcontent.opf");
                    rcont = rcont.Replace("content=\"\"", "content=\"" + kfiWidth.ToString() + "x" + kfiHeight.ToString() + "\"");
                    rcont = rcont.Replace("<dc:title></dc:title>", "<dc:title>" + doc_title + "</dc:title>");
                    rcont = rcont.Replace("</dc:author>", doc_author + "</dc:author>");
                    mdyDate = DateTime.Today.ToString("yyyy");
                    rcont = rcont.Replace("<dc:date></dc:date>", "<dc:date>" + mdyDate + "</dc:date>");
                    rcont = rcont.Replace("</metadata>", "</metadata>" + contX);
                    File.WriteAllText(kfwrkDir + "\\content.opf", rcont);
                }
                #endregion

                #endregion



                #region save as image


                for (int pageCount = 1; pageCount <= pdfDocument.Pages.Count; pageCount++)
                {
                    probar_update("Image extract...", pdfDocument.Pages.Count, pageCount);


                    using (FileStream imageStream = new FileStream(imgDir + "\\page" + pageCount + ".jpg", FileMode.Create))
                    {
                        Resolution resolution = new Resolution(imgDPI);
                        JpegDevice jpegDevice = new JpegDevice(resolution, 100);
                        jpegDevice.Process(pdfDocument.Pages[pageCount], imageStream);
                        imageStream.Close();
                    }
                }

                //cover image
                try
                {
                    if (File.Exists(cPDF.usrfolder + "\\cover.jpg"))
                    {
                        File.Copy(cPDF.usrfolder + "\\cover.jpg", imgDir + "\\cover.jpg", true);
                        File.Delete(cPDF.usrfolder + "\\cover.jpg");
                    }
                    else
                    {

                        File.Copy(imgDir + "\\page1.jpg", imgDir + "\\cover.jpg", true);
                    }

                }
                catch { }


                #region kf image

                if (cPDF.b_kf)
                {
                    for (int pageCount = 1; pageCount <= pdfDocument.Pages.Count; pageCount++)
                    {
                        probar_update("Image extract...", pdfDocument.Pages.Count, pageCount);


                        using (FileStream imageStream = new FileStream(kfimgDir + "\\page" + pageCount + ".jpg", FileMode.Create))
                        {
                            Resolution resolution = new Resolution(96);
                            JpegDevice jpegDevice = new JpegDevice(resolution, 100);
                            jpegDevice.Process(pdfDocument.Pages[pageCount], imageStream);
                            imageStream.Close();
                        }
                    }

                    //cover image
                    try
                    {
                        if (File.Exists(cPDF.usrfolder + "\\cover_kf.jpg"))
                        {
                            File.Copy(cPDF.usrfolder + "\\cover_kf.jpg", kfimgDir + "\\cover.jpg", true);
                            File.Delete(cPDF.usrfolder + "\\cover_kf.jpg");
                        }
                        else
                        {

                            File.Copy(kfimgDir + "\\page1.jpg", kfimgDir + "\\cover.jpg", true);
                        }

                    }
                    catch { }

                }
                #endregion


                #endregion



                probar_update("Extract fonts...", 0, 0);
                #region extract font

                //File.Copy(inPDFpath, styDir + "\\" + PDFfName + ".pdf", true);

                //if (File.Exists(styDir + "\\out.txt"))
                //{
                //    try { File.Delete(styDir + "\\out.txt"); }
                //    catch { }
                //}

                //DateTimeFormatInfo sv = CultureInfo.CurrentCulture.DateTimeFormat;
                //string sdf = sv.ShortDatePattern;
                //string curDate_val = DateTime.Today.ToString(sdf);
                //DateTime licDate = DateTime.ParseExact("04-01-2012", "MM-dd-yyyy", System.Globalization.CultureInfo.InvariantCulture);
                //string ulicDate = licDate.ToString(sdf);
                //gCls.setdate_update(ulicDate);

                Directory.SetCurrentDirectory(styDir);
                ProcessStartInfo psInfo = new ProcessStartInfo();
                psInfo.CreateNoWindow = true;
                psInfo.UseShellExecute = false;
                psInfo.RedirectStandardOutput = true;
                psInfo.WindowStyle = ProcessWindowStyle.Hidden;
                //psInfo.FileName = "pdfextract";
                psInfo.FileName = Application.StartupPath + "\\pdfex\\pdfextract.exe";
               // psInfo.Arguments = "-lf -x \"" + PDFfName + ".pdf\" -lk " + gCls.pdfex_key;
                psInfo.Arguments = "-$ "+ gCls.pdfex_key + "\"" + inPDFpath + "\" \"" + styDir + "\"";
                try
                {
                    string outE = "";
                    using (Process exeProcess = Process.Start(psInfo))
                    {
                        outE = exeProcess.StandardOutput.ReadToEnd();
                        exeProcess.WaitForExit();
                    }

                    //File.WriteAllText(styDir + "\\out.txt", outE);
                }
                catch (Exception erd)
                {
                    curErrLog += erd.Message.ToString();
                }
                //update curdate
                //gCls.setdate_update(curDate_val);


                //try { File.Delete(styDir + "\\" + PDFfName + ".pdf"); }
                //catch { }



                //rename font filename
                #region rename font filename

                string[] styFontcol = Directory.GetFiles(styDir, "*.*");
                foreach (string sfx in styFontcol)
                {
                    string fntName = Path.GetFileName(sfx);
                    fntName = Regex.Replace(fntName, "(.+?)\\+(.+?)", "$2");
                    fntName = Regex.Replace(fntName, "(.+?)\\-(\\d+)\\.(.+?)", "$1.$3");
                    File.Move(sfx, styDir + "\\" + fntName);
                }

                #endregion

                #region font copy to kf dir
                if (cPDF.b_kf)
                {
                    string[] kfFonts = Directory.GetFiles(styDir, "*.*");
                    foreach (string kfx in kfFonts)
                    {
                        File.Copy(kfx, kfstyDir + "\\" + Path.GetFileName(kfx), true);
                    }
                }
                  

                #endregion

                #region read out.txt
                //string[] fntTxt = File.ReadAllLines(styDir + "\\out.txt");
                //foreach (string fnt in fntTxt)
                //{
                //    try
                //    {
                //        string[] fstr = fnt.Split(',');
                //        string fntName = "";
                //        if (fstr[2].IndexOf("+") != -1)
                //        {
                //            fntName = fstr[2].Split('+')[1];
                //        }
                //        else { fntName = fstr[2]; }                        

                //        fntName = fntName.Replace("\"", "");
                //        string curfntFileName = "";
                //        if (fstr.Length == 10)
                //        {
                //            curfntFileName = fstr[9];
                //            fntName += " " + fstr[3].Replace("\"", "");
                //        }
                //        else { curfntFileName = fstr[8]; }
                //        if (File.Exists(styDir + "\\" + curfntFileName))
                //        {
                //            string fExt = Path.GetExtension(curfntFileName);
                //            if (!File.Exists(styDir + "\\" + fntName + fExt))
                //            {
                //                if (cPDF.b_kf)
                //                {
                //                    File.Copy(styDir + "\\" + curfntFileName, kfstyDir + "\\" + fntName + fExt);
                //                }
                //                File.Move(styDir + "\\" + curfntFileName, styDir + "\\" + fntName + fExt);

                //            }
                //            else { File.Delete(styDir + "\\" + curfntFileName); }
                //        }
                //    }
                //    catch { }
                //}

                //try
                //{
                //    File.Delete(styDir + "\\out.txt");
                //}
                //catch { }

                #endregion

                //font conversion
                #region font convertion

                string[] cFontsFile = Directory.GetFiles(styDir, "*.cff");

                foreach (string cf in cFontsFile)
                {

                    try
                    {
                        string cFilename = Path.GetFileNameWithoutExtension(cf);
                        string cffPath = Application.StartupPath + "\\ping\\ping.exe";
                        string otfPath = Application.StartupPath + "\\makeots\\makeotf.exe";
                        psInfo.CreateNoWindow = true;
                        psInfo.UseShellExecute = false;

                        psInfo.RedirectStandardOutput = true;
                        psInfo.WindowStyle = ProcessWindowStyle.Hidden;
                        psInfo.FileName = cffPath;
                        Directory.SetCurrentDirectory(Application.StartupPath + "\\ping");
                        psInfo.Arguments = "-a \"" + cf + "\" \"" + styDir + "\\" + cFilename + ".pfa\"";
                        try
                        {
                            using (Process exeProcess = Process.Start(psInfo))
                            {
                                exeProcess.WaitForExit();
                            }
                        }
                        catch (Exception erd)
                        {
                            curErrLog += erd.Message.ToString();
                        }

                        #region for kf font conversion
                        psInfo.Arguments = "-a \"" + cf + "\" \"" + kfstyDir + "\\" + cFilename + ".pfa\"";
                        if (cPDF.b_kf)
                        {
                            try
                            {
                                using (Process exeProcess = Process.Start(psInfo))
                                {
                                    exeProcess.WaitForExit();
                                }
                            }
                            catch { }
                        }
                        #endregion

                        //convert pfa to otf
                        if (File.Exists(styDir + "\\" + cFilename + ".pfa"))
                        {
                            psInfo.FileName = otfPath;
                            Directory.SetCurrentDirectory(Application.StartupPath + "\\makeots");
                            string inPfaFile = styDir + "\\" + cFilename + ".pfa";
                            string outOtffile = styDir + "\\" + cFilename + ".otf";
                            psInfo.Arguments = "-f \"" + inPfaFile + "\" -o \"" + outOtffile + "\"";
                            try
                            {
                                using (Process exeProcess = Process.Start(psInfo))
                                {
                                    exeProcess.WaitForExit();
                                }
                            }
                            catch (Exception erd)
                            {
                                curErrLog += erd.Message.ToString();
                            }

                        }

                        #region for kf font conversion
                        if (cPDF.b_kf)
                        {
                            string inPfaFile = kfstyDir + "\\" + cFilename + ".pfa";
                            string outOtffile = kfstyDir + "\\" + cFilename + ".otf";
                            psInfo.Arguments = "-f \"" + inPfaFile + "\" -o \"" + outOtffile + "\"";
                            try
                            {
                                using (Process exeProcess = Process.Start(psInfo))
                                {
                                    exeProcess.WaitForExit();
                                }
                            }
                            catch { }
                        }
                        #endregion


                        Directory.SetCurrentDirectory(Application.StartupPath);

                        try
                        {
                            File.Delete(cf); File.Delete(styDir + "\\" + cFilename + ".pfa");
                            if (cPDF.b_kf) { File.Delete(kfstyDir + "\\" + cFilename + ".pfa"); }
                        }
                        catch { }


                    }
                    catch (Exception erd)
                    {
                        curErrLog += erd.Message.ToString();
                    }

                }

                #endregion
                
                #endregion

               



                probar_update("Style apply...", 0, 0);
                #region page css
                int iWidth = 0;
                int iHeight = 0;

                
                #region picture org width & height
                try
                {
                    System.Drawing.Image imgBox = System.Drawing.Image.FromFile(imgDir + "\\page1.jpg");
                    iWidth = imgBox.Size.Width;
                    iHeight = imgBox.Size.Height;
                    imgBox.Dispose();

                    if (cPDF.b_kf)
                    {
                        System.Drawing.Image kimgBox = System.Drawing.Image.FromFile(kfimgDir + "\\page1.jpg");
                        kfiWidth = kimgBox.Size.Width;
                        kfiHeight = kimgBox.Size.Height;
                        kimgBox.Dispose();
                    }
                }
                catch { }
                #endregion

                //font info
                string fontTxt = "";
                string kffontTxt = "";
                ArrayList tmpFontList = new ArrayList();
                for (int ft = 0; ft < fntCollection.Count; ft++)
                {
                    fInfo ftx = (fInfo)fntCollection[ft];
                    bool tmpFound = false;
                    for (int tf = 0; tf < tmpFontList.Count; tf++)
                    {
                        if (tmpFontList[tf].ToString() == ftx.f_fontfamily)
                        {
                            tmpFound = true;
                        }
                    }

                    if (tmpFound == false)
                    {
                        fontTxt += "@font-face\n{\nfont-family: '" + ftx.f_fontfamily + "';\n";
                        kffontTxt += "@font-face\n{\nfont-family: '" + ftx.f_fontfamily + "';\n";

                        if (File.Exists(styDir + "\\" + ftx.f_fontfamily + ".ttf"))
                        {
                            fontTxt += "src:url(" + ftx.f_fontfamily + ".ttf);\n";
                            kffontTxt += "src:url(" + ftx.f_fontfamily + ".ttf);\n";

                        }
                        else if (File.Exists(styDir + "\\" + ftx.f_fontfamily + ".otf"))
                        {
                            fontTxt += "src:url(" + ftx.f_fontfamily + ".otf);\n";
                            kffontTxt += "src:url(" + ftx.f_fontfamily + ".otf);\n";
                        }
                        fontTxt += "}\n";
                        kffontTxt += "}\n";

                        tmpFontList.Add(ftx.f_fontfamily);
                    }
                    //resize
                    string re_fntSize = ((Convert.ToInt32(Math.Round(Convert.ToDouble(ftx.f_fontsize))) * imgDPI) / 72).ToString();
                    string kfre_fntSize = ((Convert.ToInt32(Math.Round(Convert.ToDouble(ftx.f_fontsize))) * 96) / 72).ToString();

                    fontTxt += "." + ftx.f_fontid + " {\n";
                    fontTxt += "font-size: " + re_fntSize + "px;\n";
                    fontTxt += "font-family: " + ftx.f_fontfamily + ";\n";
                    fontTxt += "color: " + ftx.f_color + ";\n";
                    fontTxt += "}\n";                    

                    kffontTxt += "." + ftx.f_fontid + " {\n";
                    kffontTxt += "font-size: " + kfre_fntSize + "px;\n";
                    kffontTxt += "font-family: " + ftx.f_fontfamily + ";\n";
                    kffontTxt += "color: " + ftx.f_color + ";\n";
                    kffontTxt += "}\n";

                }

                fontTxt += ".-epub-media-overlay-active{\n";
                fontTxt += "background-color: yellow;\n";
                fontTxt += "}\n";

                //add font into opf
                #region add font list into content.opf
                string flst = "";
                for (int ftx = 0; ftx < tmpFontList.Count; ftx++)
                {
                    if (File.Exists(styDir + "\\" + tmpFontList[ftx].ToString() + ".ttf"))
                    {
                        flst += "<item id=\"font" + (ftx + 1).ToString() + "\" href=\"styles/" + tmpFontList[ftx].ToString() + ".ttf\" media-type=\"application/x-font-ttf\"/>\n";
                    }
                    else if (File.Exists(styDir + "\\" + tmpFontList[ftx].ToString() + ".otf"))
                    {
                        flst += "<item id=\"font" + (ftx + 1).ToString() + "\" href=\"styles/" + tmpFontList[ftx].ToString() + ".otf\" media-type=\"application/x-font-otf\"/>\n";
                    }
                }
                string fntOpf = File.ReadAllText(oebDir + "\\content.opf");
                fntOpf = fntOpf.Replace("</manifest>", flst + "</manifest>");
                File.WriteAllText(oebDir + "\\content.opf", fntOpf);

                #region write kf content.opf
                if (cPDF.b_kf)
                {
                    fntOpf = File.ReadAllText(kfwrkDir + "\\content.opf");
                    fntOpf = fntOpf.Replace("</manifest>", flst.Replace("href=\"styles/", "href=\"css/") + "</manifest>");
                    File.WriteAllText(kfwrkDir + "\\content.opf", fntOpf);
                }
                #endregion

                #endregion

                fontTxt += "body\n{\n";
                //fontTxt += "width: " + iWidth.ToString() + "px;\n";
                //fontTxt += "height: " + iHeight.ToString() + "px;\n";
                fontTxt += "margin: 0 0 0 0px;\n";
                fontTxt += "padding: 0 0 0 0px;\n";
                fontTxt += "}\n";

                fontTxt += ".leftspread {\n";
                //fontTxt += "width: " + iWidth.ToString() + "px;\n";
                //fontTxt += "height: " + iHeight.ToString() + "px;\n";
                fontTxt += "padding: 0em;\n";
                fontTxt += "margin-right: 0em;\n";
                fontTxt += "}\n";

                fontTxt += ".rightspread {\n";
                fontTxt += "width: " + iWidth.ToString() + "px;\n";
                fontTxt += "height: " + iHeight.ToString() + "px;\n";
                fontTxt += "padding: 0em;\n";
                fontTxt += "margin-left: 0em;\n";
                fontTxt += "}\n";

                #region kffont
                kffontTxt += "body\n{\n";
                //kffontTxt += "width: " + iWidth.ToString() + "px;\n";
                //kffontTxt += "height: " + iHeight.ToString() + "px;\n";
                kffontTxt += "margin: 0 0 0 0px;\n";
                kffontTxt += "padding: 0 0 0 0px;\n";
                kffontTxt += "}\n";

                kffontTxt += ".leftspread {\n";
                //kffontTxt += "width: " + iWidth.ToString() + "px;\n";
                //kffontTxt += "height: " + iHeight.ToString() + "px;\n";
                kffontTxt += "padding: 0em;\n";
                kffontTxt += "margin-right: 0em;\n";
                kffontTxt += "}\n";

                kffontTxt += ".rightspread {\n";
                kffontTxt += "width: " + iWidth.ToString() + "px;\n";
                kffontTxt += "height: " + iHeight.ToString() + "px;\n";
                kffontTxt += "padding: 0em;\n";
                kffontTxt += "margin-left: 0em;\n";
                kffontTxt += "}\n";

                #endregion

                File.WriteAllText(styDir + "\\page.css", fontTxt);

                if (cPDF.b_kf)
                {
                    File.WriteAllText(kfstyDir + "\\page.css", kffontTxt);
                }

                #endregion



                #region create each page css and html
                //create cover html

                gCls.get_image_size(imgDir + "\\cover.jpg", ref iWidth, ref iHeight);
                gCls.get_image_size(kfimgDir + "\\cover.jpg", ref kfiWidth, ref kfiHeight);

                string converTxt = "<!DOCTYPE html>\n";
                converTxt += "<html xmlns=\"http://www.w3.org/1999/xhtml\">\n";
                converTxt += "<head>\n";
                converTxt += "<title>" + doc_title + "</title>\n";
                converTxt += "<meta name=\"viewport\" content=\"width=" + iWidth.ToString() + ", height=" + iHeight.ToString() + "\"/>\n";
                converTxt += "<link rel=\"stylesheet\" type=\"text/css\" href=\"../styles/page.css\" />\n";
                converTxt += "</head><body><div class=\"rightspread\"><div class=\"one\"><img src=\"../images/cover.jpg\" width=\"" + iWidth.ToString() + "\" height=\"" + iHeight.ToString() + "\" alt=\"\" /></div></div></body></html>\n";
                File.WriteAllText(htmlDir + "\\cover.html", converTxt);



                if (cPDF.b_kf)
                {
                    converTxt = "<!DOCTYPE html>\n";
                    converTxt += "<html xmlns=\"http://www.w3.org/1999/xhtml\">\n";
                    converTxt += "<head>\n";
                    converTxt += "<title>" + doc_title + "</title>\n";
                    converTxt += "<meta name=\"viewport\" content=\"width=" + kfiWidth.ToString() + ", height=" + kfiHeight.ToString() + "\"/>\n";
                    converTxt += "<link rel=\"stylesheet\" type=\"text/css\" href=\"../styles/page.css\" />\n";
                    converTxt += "</head><body><div class=\"rightspread\"><div class=\"one\"><img src=\"../image/cover.jpg\" width=\"" + kfiWidth.ToString() + "\" height=\"" + kfiHeight.ToString() + "\" alt=\"\" /></div></div></body></html>\n";
                    File.WriteAllText(kfhtmlDir + "\\cover.html", converTxt);
                }

                //toc.html

                //string tocTxt = "<html xmlns=\"http://www.w3.org/1999/xhtml\" xmlns:epub=\"http://www.idpf.org/2007/ops\"><head><title>" + doc_title + "</title></head><body>\n";
                //tocTxt += "<section epub:type=\"frontmatter toc\"><header><h1>Contents</h1></header><nav epub:type=\"toc\" id=\"toc\"><ol><li id=\"chapter_001\"> <a href=\"cover.html\">Home</a></li></ol></nav></section></body></html>";
                //File.WriteAllText(htmlDir + "\\toc.html", tocTxt);

                //==



                for (int pg = 1; pg < pdfDocument.Pages.Count + 1; pg++)
                {

                    probar_update("ePub xhtml creation...", pdfDocument.Pages.Count, pg);



                    ArrayList pLs = (ArrayList)pageCollection[pg - 1];

                    string pageCSS = "";
                    string kfpageCSS = "";
                    string pageHtml = "";
                    string kfpageHtml = "";

                    gCls.get_image_size(imgDir + "\\page" + pg.ToString() + ".jpg", ref iWidth, ref iHeight);
                    gCls.get_image_size(kfimgDir + "\\page" + pg.ToString() + ".jpg", ref kfiWidth, ref kfiHeight);

                    pageHtml += "<!DOCTYPE HTML>\n";
                    pageHtml += "<html xmlns=\"http://www.w3.org/1999/xhtml\" xmlns:ibooks=\"http://apple.com/ibooks/html-extensions\" xmlns:epub=\"http://www.idpf.org/2007/ops\">\n";
                    pageHtml += "<head>\n";
                    pageHtml += "<title>" + doc_title + " : Page" + pg.ToString() + "</title>\n";
                    pageHtml += "<link rel=\"stylesheet\" type=\"text/css\" href=\"../styles/page.css\" />\n";
                    pageHtml += "<link rel=\"stylesheet\" type=\"text/css\" href=\"../styles/page" + pg.ToString() + ".css\" />\n";

                    

                    pageHtml += "<meta name=\"viewport\" content=\"width=" + iWidth.ToString() + ", height=" + iHeight.ToString() + "\"/>\n";
                    pageHtml += "</head>\n";
                    pageHtml += "<body>\n";
                    pageHtml += "<div class=\"leftspread\">\n";
                    pageHtml += "<img src=\"../images/page" + pg.ToString() + ".jpg\" width=\"" + iWidth.ToString() + "\" height=\"" + iHeight.ToString() + "\" alt=\"\"/>\n";
                    pageHtml += "<div>\n";

                    if (cPDF.b_kf)
                    {
                        kfpageHtml += "<!DOCTYPE HTML>\n";
                        kfpageHtml += "<html xmlns=\"http://www.w3.org/1999/xhtml\" xmlns:ibooks=\"http://apple.com/ibooks/html-extensions\" xmlns:epub=\"http://www.idpf.org/2007/ops\">\n";
                        kfpageHtml += "<head>\n";
                        kfpageHtml += "<title>" + doc_title + " : Page" + pg.ToString() + "</title>\n";
                        kfpageHtml += "<link rel=\"stylesheet\" type=\"text/css\" href=\"../css/page.css\" />\n";
                        kfpageHtml += "<link rel=\"stylesheet\" type=\"text/css\" href=\"../css/page" + pg.ToString() + ".css\" />\n";
                        kfpageHtml += "<meta name=\"viewport\" content=\"width=" + kfiWidth.ToString() + ", height=" + kfiHeight.ToString() + "\"/>\n";
                        kfpageHtml += "</head>\n";
                        kfpageHtml += "<body>\n";
                        kfpageHtml += "<div class=\"leftspread\" style=\"background-repeat: no-repeat; background-image: url('../image/page" + pg.ToString() + ".jpg'); width:" + kfiWidth.ToString() + "px; height:" + kfiHeight.ToString() + "px;\" >\n";
                        kfpageHtml += "</div>\n";

                    }

                   
                    htLineTxt = "";
                    ArrayList divList = new ArrayList();
                    for (int txi = 0; txi < pLs.Count; txi++)
                    {
                        tInfo txInfo = (tInfo)pLs[txi];
                        int divID = txi + 1;
                        string p_top = Math.Round(double.Parse(txInfo.p_top)).ToString();
                        string p_left = Math.Round(double.Parse(txInfo.p_left)).ToString();

                        

                        string kf_top = ((Convert.ToInt32(Convert.ToDouble(p_top)) * 96) / 72).ToString();
                        string kf_left = ((Convert.ToInt32(Convert.ToDouble(p_left)) * 96) / 72).ToString();

                        p_top = ((Convert.ToInt32(Convert.ToDouble(p_top)) * imgDPI) / 72).ToString();
                        p_left = ((Convert.ToInt32(Convert.ToDouble(p_left)) * imgDPI) / 72).ToString();

                        


                        string p_text = txInfo.p_text;
                        p_text = p_text.Replace("&", "&amp;");
                        p_text = p_text.Replace("<", "&lt;");
                        p_text = p_text.Replace(">", "&gt;");

                        
                        
                        if (chk_audio.Checked && check_audio_custom_page(int.Parse(pg.ToString())))
                        {
                            htLineTxt += p_text + "\n";
                            divList.Add("<div id=\"d" + divID.ToString() + "\" class=\"" + txInfo.p_fontid + " text" + divID.ToString() + "\">");
                        }

                        p_text = p_text.Replace(" ", "&#160;");
                        p_text = gCls.Text2HexaConversion(p_text);


                        if (!chk_audio.Checked || check_audio_custom_page(int.Parse(pg.ToString())) == false)
                        {
                            pageHtml += "<div id=\"d" + divID.ToString() + "\" class=\"" + txInfo.p_fontid + " text" + divID.ToString() + "\">";                            
                            pageHtml += p_text;
                            pageHtml += "</div>\n";
                            
                        }

                        if (cPDF.b_kf)
                        {
                            kfpageHtml += "<div id=\"d" + divID.ToString() + "\" class=\"" + txInfo.p_fontid + " text" + divID.ToString() + "\">" + p_text + "</div>\n";
                        }
                        pageCSS += ".text" + divID.ToString() + " { position:absolute;top:" + p_top + "px;left:" + p_left + "px; }\n";

                        if (cPDF.b_kf)
                        {
                            kfpageCSS += ".text" + divID.ToString() + " { position:absolute;top:" + kf_top + "px;left:" + kf_left + "px; }\n";
                        }
                    }
                    

                    #region read aloud
                    if (chk_audio.Checked && check_audio_custom_page(int.Parse(pg.ToString())))
                    {

                        string readLineTxt = htLineTxt.Replace("-\n", "");
                        readLineTxt = readLineTxt.Replace("\n", " ");
                        readLineTxt = readLineTxt.Replace("&#160;", " ");
                        readLineTxt = readLineTxt.Replace("&lt;", "<");
                        readLineTxt = readLineTxt.Replace("&gt;", ">");
                        readLineTxt = readLineTxt.Replace("&amp;", "&");

                        read_aloud_html(readLineTxt,audioDir,htmlDir,pg.ToString(),voiceName);

                        htLineTxt = gCls.Text2HexaConversion(htLineTxt);
                        htLineTxt = htLineTxt.Replace("<span id=", "<span&#161;id=");
                        htLineTxt = htLineTxt.Replace(" ", "&#160;");
                        htLineTxt = htLineTxt.Replace("<span&#161;id=", "<span id=");

                        string[] hLinecol = htLineTxt.Split('\n');                        
                        for (int h = 0; h < divList.Count; h++) {
                            pageHtml += divList[h].ToString() + hLinecol[h] + "</div>\n";
                        }
                    }
                    #endregion

                    #region check attachement media
                    if (attList.Count > 0) {
                        for (int s = 0; s < attList.Count; s++) {
                            string[] gatList = attList[s].ToString().Split('|');
                            string pgString = gatList[0];
                            string attName = gatList[1];
                            string atX = gatList[2];
                            string atY = gatList[3];
                            string atWidth = gatList[4];
                            string atHeight = gatList[5];
                            string mtType = gatList[6];
                            if (pgString == pg.ToString()) {
                                string gExtName = Path.GetExtension(attName);
                                gExtName = gExtName.Replace(".", "");

                                if (mtType == "audio") {
                                    pageHtml += "<audio controls=\"controls\" style=\"position:absolute; left:" + atX + "px; top:" + atY + "px;\"><source src=\"../images/" + attName + "\" type=\"audio/" + gExtName + "\" />  </audio>";
                                }
                                else if (mtType == "video") {
                                    pageHtml += "<video controls=\"controls\" style=\"width:" + atWidth + "px; height:" + atHeight + "px; position:absolute; left:" + atX + "px; top:" + atY + "px;\"><source src=\"../images/" + attName + "\" type=\"video/" + gExtName + "\" />  </video>";
                                }                            
                            }                        
                        }                    
                    }
                    #endregion

                    pageHtml += "</div></div>\n";                    
                    
                    pageHtml += "</body></html>\n";

                    kfpageHtml += "</body></html>\n";

                    File.WriteAllText(styDir + "\\page" + pg.ToString() + ".css", pageCSS);
                    File.WriteAllText(htmlDir + "\\page" + pg.ToString() + ".html", pageHtml);

                    if (cPDF.b_kf)
                    {
                        File.WriteAllText(kfstyDir + "\\page" + pg.ToString() + ".css", kfpageCSS);
                        File.WriteAllText(kfhtmlDir + "\\page" + pg.ToString() + ".html", kfpageHtml);
                    }
                }
                
                #endregion

                if (chk_audio.Checked)
                {
                   
                    convert_mp3(audioDir);
                }

                File.Delete(inPDFpath);
                try
                {
                    if (File.Exists(cPDF.usrfolder + "\\" + cPDF.b_filename + ".epub"))
                    {
                        File.Delete(cPDF.usrfolder + "\\" + cPDF.b_filename + ".epub");
                    }
                }
                catch { }

                inPDFPageCount = pdfDocument.Pages.Count;


                probar_update("ePub package creation...", 0, 0);

                string[] content = { wrkDir + "\\mimetype", wrkDir + "\\OEBPS", wrkDir + "\\META-INF" };

                //epub package
                #region epub package

                Directory.SetCurrentDirectory(wrkDir);
                psInfo.CreateNoWindow = true;
                psInfo.UseShellExecute = false;
                psInfo.RedirectStandardOutput = true;
                psInfo.WindowStyle = ProcessWindowStyle.Hidden;
                psInfo.FileName = "ezip.exe";
                psInfo.Arguments = " -Xr9D \"" + cPDF.usrfolder + "\\" + cPDF.b_filename + ".epub\" mimetype *";
                //psInfo.Arguments = "-Xr9D " + cPDF.b_filename + ".epub mimetype *";
                try
                {

                    using (Process exeProcess = Process.Start(psInfo))
                    {
                        string outE = exeProcess.StandardOutput.ReadToEnd();
                        exeProcess.WaitForExit();
                    }
                }
                catch (Exception erd)
                {
                    curErrLog += erd.Message.ToString();
                }
                Directory.SetCurrentDirectory(cPDF.usrfolder);

                //Ionic.Zip.ZipFile ezip = new Ionic.Zip.ZipFile();
                //ezip.AddEntry("mimetype", "application/epub+zip").CompressionLevel = Ionic.Zlib.CompressionLevel.None;
                //ezip.AddDirectory(wrkDir + "\\OEBPS", "OEBPS");
                //ezip.AddDirectory(wrkDir + "\\META-INF", "META-INF");
                //ezip.Save(cPDF.usrfolder + "\\" + cPDF.b_filename + ".epub");

                if (File.Exists(cPDF.usrfolder + "\\" + cPDF.b_filename + ".epub"))
                {
                    try { Directory.Delete(wrkDir, true); }
                    catch { }

                }
                #endregion

                //kf package kindlegen
                #region kindlegen package
                if (cPDF.b_kf) {
                    if (File.Exists(kfwrkDir + "\\content.opf")) {
                        psInfo.FileName = "kindlegen.exe";
                        Directory.SetCurrentDirectory(Application.StartupPath + "\\kindlegen");
                        
                        string kfcontent_opf = kfwrkDir + "\\content.opf";
                        psInfo.Arguments = "\""  + kfcontent_opf + "\"";
                        try
                        {
                            using (Process exeProcess = Process.Start(psInfo))
                            {
                                exeProcess.WaitForExit();
                            }
                        }
                        catch (Exception erd)
                        {
                            curErrLog += erd.Message.ToString();
                        }

                    }

                    if (File.Exists(kfwrkDir + "\\content.mobi")) {
                        File.Copy(kfwrkDir + "\\content.mobi", cPDF.usrfolder + "\\" + cPDF.b_filename + ".mobi",true);
                        try { Directory.Delete(kfwrkDir, true); }
                        catch { }
                    }
                   
                }
                #endregion

                probar_update("Completed", 0, 0);

            }
            catch (Exception erd)
            {
                curErrLog += erd.Message.ToString();
            }

        }


        public void read_aloud_html(string inHtml,string audir,string htmlDir,string pgNum,string voicename){
            try
            {


                System.Speech.Synthesis.SpeechSynthesizer l_spv = new System.Speech.Synthesis.SpeechSynthesizer();
                l_spv.SpeakCompleted += new EventHandler<System.Speech.Synthesis.SpeakCompletedEventArgs>(spv_SpeakCompleted);
                l_spv.SpeakProgress += new EventHandler<System.Speech.Synthesis.SpeakProgressEventArgs>(spv_SpeakProgress);
               
                l_spv.SelectVoice(voicename);
                
                l_spv.Rate = -2;
                l_spv.SetOutputToWaveFile(audir + "\\audio" + pgNum + ".wav");
                audioComplete = false;
                spanTxt = "";
                smilTxt = "";
                last_time = "";
                curFullTxt = "";
                curFullTxt = inHtml;
                spanID = 1;
                spanPoint = 0;
                curPageHtml = "page" + pgNum + ".html";
                curaudioFile = "audio" + pgNum + ".mp3";

                l_spv.SpeakAsync(inHtml);
                while (audioComplete == false)
                {
                    Application.DoEvents();
                }

                string wSmil = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n";             
                wSmil += "<smil xmlns=\"http://www.w3.org/ns/SMIL\" version=\"3.0\" profile=\"http://www.idpf.org/epub/30/profile/content/\">\n";             
                wSmil += "<body>\n";
                wSmil += smilTxt;
                wSmil += "\n</body></smil>\n";
                File.WriteAllText(htmlDir + "\\page" + pgNum + ".smil", wSmil);

                l_spv.Dispose();
                
            }
            catch {
                
            }


        }

        private void spv_SpeakCompleted(object sender, System.Speech.Synthesis.SpeakCompletedEventArgs e)
        {
            audioComplete = true;
            string outsmile = smilTxt;
            outsmile = outsmile.Replace("[etime/]", last_time);
            smilTxt = outsmile;

        }
        private void spv_SpeakProgress(object sender, System.Speech.Synthesis.SpeakProgressEventArgs e)
        {


            TimeSpan df = e.AudioPosition;

            //double cPT = double.Parse(Math.Round(df.TotalSeconds).ToString() + "." +  Math.Round(df.TotalMilliseconds).ToString());
            double cPT = df.TotalSeconds;            
            double xPT = (cPT * 30) / 100;
            xPT = xPT - (xPT * 13 / 136);
            cPT = (cPT - xPT) ;

            //string aTime = df.Seconds.ToString() + "." + df.Milliseconds.ToString() + "s";
            //string caTime = df.Seconds.ToString() + "." + (df.Milliseconds - 2).ToString() + "s";
            string aTime = cPT.ToString() + "s";
            string caTime = cPT.ToString() + "s";

            last_time = (cPT + xPT).ToString() + "s";
            

            string spTxt = "";
            #region pro text
            int sPoint = e.CharacterPosition;

            if (curFullTxt.IndexOf(" ", sPoint) != -1)
            {
                int ePoint = curFullTxt.IndexOf(" ", sPoint);
                try
                {
                    if (curFullTxt.Substring(sPoint - 1, 1) == "\"")
                    {
                        sPoint = sPoint - 1;
                    }
                }
                catch { }
                spTxt = curFullTxt.Substring(sPoint, ePoint - sPoint);
                
            }
            else { spTxt = e.Text; }
            #endregion

            spTxt = spTxt.Replace("<", "&lt;");
            spTxt = spTxt.Replace(">", "&gt;");
            spTxt = spTxt.Replace("&", "&amp;");
           
            Regex wFnx = new Regex("\\s*(.+?)\\s", RegexOptions.IgnoreCase);
            string mtchTxt = "";
            MatchCollection mtchPoint = null;
            int mInt = 0;
          

            if (wFnx.Match(htLineTxt, spanPoint).Success) {
                mtchPoint = wFnx.Matches(htLineTxt,spanPoint);
                try {
                    if (mtchPoint[0].Groups[1].Value == spTxt) {
                        mtchTxt = mtchPoint[0].Groups[1].Value;
                        mInt = 0;
                    }
                    else if (mtchPoint[1].Groups[1].Value == spTxt) {
                        mtchTxt = mtchPoint[1].Groups[1].Value;
                        mInt = 1;
                    }
                    else if (mtchPoint[2].Groups[1].Value == spTxt) {
                        mtchTxt = mtchPoint[2].Groups[1].Value;
                        mInt = 2;
                    }
                    else if (mtchPoint[3].Groups[1].Value == spTxt)
                    {
                        mtchTxt = mtchPoint[3].Groups[1].Value;
                        mInt = 3;
                    }
                    else if (mtchPoint[4].Groups[1].Value == spTxt)
                    {
                        mtchTxt = mtchPoint[4].Groups[1].Value;
                        mInt = 4;
                    }

                }
                catch { }
                
            }
            
            if (spTxt == mtchTxt) {
                int stPoint = mtchPoint[mInt].Groups[1].Index;
                int etPoint = stPoint + spTxt.Length;
                htLineTxt = htLineTxt.Insert(etPoint, "</span>");
                string spanidTxt = "<span id=\"w" + spanID.ToString() + "\">";
                htLineTxt = htLineTxt.Insert(stPoint, spanidTxt);
                

                spanPoint = etPoint + 7 + spanidTxt.Length;
            }            

            smilTxt = smilTxt.Replace("[etime/]", caTime);

            string pTxt = "<par id=\"par" + spanID.ToString() + "\">\n";
            pTxt += "<text src=\"" + curPageHtml + "#w" + spanID.ToString() + "\" />\n";
            pTxt += "<audio src=\"../audio/" + curaudioFile + "\" clipBegin=\"" + aTime + "\" clipEnd=\"[etime/]\"/>\n";
            pTxt += "</par>\n";
            smilTxt += pTxt;

            spanID++;

        }


        public void convert_mp3(string wavPath)
        {

            string[] wavFiles = Directory.GetFiles(wavPath, "*.wav");
            foreach (string w in wavFiles)
            {
                string sPath = Path.GetDirectoryName(w);
                string fName = Path.GetFileNameWithoutExtension(w);
                #region write mp3
                WaveStream InStr = new WaveStream(w);
                try
                {
                    Mp3Writer writer = new Mp3Writer(new FileStream(sPath + "\\" + fName + ".mp3",
                                                        FileMode.Create), InStr.Format);
                    try
                    {
                        byte[] buff = new byte[writer.OptimalBufferSize];
                        int read = 0;
                        while ((read = InStr.Read(buff, 0, buff.Length)) > 0)
                        {
                            writer.Write(buff, 0, read);
                        }
                    }
                    finally
                    {
                        writer.Close();
                    }
                }
                finally
                {
                    InStr.Close();
                }

                #endregion
                File.Delete(w);
            }

        }

        public void aspose_license_update()
        {
            try
            {
                gCls.getlicstring();
                byte[] byteArray = Encoding.ASCII.GetBytes(gCls.aspose_key);
                MemoryStream Lstr = new MemoryStream(byteArray);                 
                gCls.aspose_license.SetLicense(Lstr);
                
            }
            catch(Exception erd)  {
                gCls.show_error("Unable to load license");
                return;            
            }
        }

        private void kryptonButton1_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog fld = new OpenFileDialog();
                fld.Title = "Select PDF File";
                fld.Filter = "PDF File|*.pdf";
                fld.ShowDialog();
                if(fld.FileName != ""){
                 txt_inpdf.Text = fld.FileName;
                }
            }
            catch (Exception erd) { 
             gCls.show_error(erd.Message.ToString());
                return;
            }
        }

        private void kryptonButton3_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog fld = new OpenFileDialog();
                fld.Title = "Select cover image file";
                fld.Filter = "JPG File|*.jpg";
                fld.ShowDialog();
                if (fld.FileName != "")
                {
                    txt_epubcover.Text = fld.FileName;
                }
            }
            catch (Exception erd)
            {
                gCls.show_error(erd.Message.ToString());
                return;
            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                gCls.show_message(Application.ProductName + " " + Application.ProductVersion + "\nSend your Feedbacks to : vickypatel2020@gmail.com\n");
            }
            catch { }
        }

        private void chk_audio_CheckedChanged(object sender, EventArgs e)
        {
            if (chk_audio.Checked)
            {
                cmb_aupage.Visible = true;
                lb_aupage.Visible = true;
                cmb_voicelist.Enabled = true;
            }
            else { 
                cmb_aupage.Visible = false;
                txt_aupagetxt.Visible = false;
                lb_aupage.Visible = false;
                cmb_voicelist.Enabled = false;
            }
        }

        private void cmb_aupage_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmb_aupage.SelectedIndex == 0)
            {
                txt_aupagetxt.Visible = false;
            }
            else {
                txt_aupagetxt.Visible = true;
                }
        }

        private void chk_crop_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if (chk_crop.Checked)
                {
                    txt_crop_left.Enabled = true;
                    txt_crop_right.Enabled = true;
                    txt_crop_top.Enabled = true;
                    txt_crop_bottom.Enabled = true;
                }
                else {
                    txt_crop_left.Enabled = false;
                    txt_crop_right.Enabled = false;
                    txt_crop_top.Enabled = false;
                    txt_crop_bottom.Enabled = false;
                }

            }
            catch (Exception erd) {
                gCls.show_error(erd.Message.ToString());
                return;
            }
        }
    }
}
