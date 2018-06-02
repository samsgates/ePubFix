using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

/// <summary>
/// Summary description for Updf
/// </summary>
public class Updf
{
    
    public string usrfolder { get; set; }

    public string b_filename { get; set; }
    public string b_title { get; set; }
    public string b_author { get; set; }
    public string b_filesize { get; set; }
    public string b_resolution { get; set; }
    public string b_pagetype { get; set; }
    public bool b_kf { get; set; }

    public Updf(string i_filename, string i_title, string i_author, string i_filesize, string i_resolution, string i_usrfoler, string i_pagetype, bool i_kf)
    {
        try
        {
            

            b_filename = i_filename;
            b_title = i_title;
            b_author = i_author;
            b_filesize = i_filesize;
            b_resolution = i_resolution;
            usrfolder = i_usrfoler;
            b_pagetype = i_pagetype;
            b_kf = i_kf;
        }
        catch { }
    }
}

public class tInfo
{
    public string p_left { get; set; }
    public string p_top { get; set; }
    public string p_text { get; set; }    
    public string p_fontid { get; set; }

    public tInfo(string left, string top, string text, string fontid)
    {
        p_left = left;
        p_top = top;
        p_text = text;        
        p_fontid = fontid;
    }


}

public class fInfo
{
    public string f_fontsize { get; set; }
    public string f_fontfamily { get; set; }
    public string f_color { get; set; }
    public string f_fontid { get; set; }


    public fInfo(string fontsize, string fontfamily, string color, string fontid)
    {
        f_fontsize = fontsize;
        f_fontfamily = fontfamily;
        f_color = color;
        f_fontid = fontid;

    }


}