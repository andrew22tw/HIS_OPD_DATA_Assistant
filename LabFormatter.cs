// Lab Data Formatter v1.1.1
// Author: \u5433\u5cb3\u9716\u91ab\u5e2b (DAL93@tpech.gov.tw)
// Compile: build.bat (auto-finds csc.exe)
// Hotkeys: Ctrl+0=Settings, Ctrl+1~4=Custom slots
// Workflow: Ctrl+C report -> auto convert -> Ctrl+3 paste

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;

// ══════════ JSON mini parser (no external deps) ══════════
static class Json {
    public static Dictionary<string,object> Parse(string s) {
        int i = 0; return ParseObj(s, ref i);
    }
    static void Skip(string s, ref int i) { while(i<s.Length&&char.IsWhiteSpace(s[i]))i++; }
    static string ParseStr(string s, ref int i) {
        i++; var sb=new System.Text.StringBuilder();
        while(i<s.Length&&s[i]!='"'){if(s[i]=='\\'){i++;if(i<s.Length){
            if(s[i]=='n')sb.Append('\n');else if(s[i]=='t')sb.Append('\t');
            else if(s[i]=='u'){sb.Append((char)Convert.ToInt32(s.Substring(i+1,4),16));i+=4;}
            else sb.Append(s[i]);}}else sb.Append(s[i]);i++;}
        i++; return sb.ToString();
    }
    static object ParseVal(string s, ref int i) {
        Skip(s,ref i); if(i>=s.Length)return null;
        if(s[i]=='"')return ParseStr(s,ref i);
        if(s[i]=='{')return ParseObj(s,ref i);
        if(s[i]=='[')return ParseArr(s,ref i);
        if(s[i]=='t'){i+=4;return true;} if(s[i]=='f'){i+=5;return false;}
        if(s[i]=='n'){i+=4;return null;}
        int st=i; while(i<s.Length&&(char.IsDigit(s[i])||s[i]=='.'||s[i]=='-'||s[i]=='e'||s[i]=='E'||s[i]=='+'))i++;
        return s.Substring(st,i-st);
    }
    static Dictionary<string,object> ParseObj(string s, ref int i) {
        var d=new Dictionary<string,object>(); i++; Skip(s,ref i);
        while(i<s.Length&&s[i]!='}'){Skip(s,ref i);if(s[i]=='}')break;
            var k=ParseStr(s,ref i);Skip(s,ref i);i++;Skip(s,ref i);
            d[k]=ParseVal(s,ref i);Skip(s,ref i);if(i<s.Length&&s[i]==',')i++;}
        i++; return d;
    }
    static List<object> ParseArr(string s, ref int i) {
        var a=new List<object>(); i++; Skip(s,ref i);
        while(i<s.Length&&s[i]!=']'){if(s[i]==']')break;
            a.Add(ParseVal(s,ref i));Skip(s,ref i);if(i<s.Length&&s[i]==',')i++;}
        i++; return a;
    }
    public static string Encode(object o, int indent) {
        if(o==null) return "null";
        if(o is string){var str=(string)o; return "\""+str.Replace("\\","\\\\").Replace("\"","\\\"").Replace("\n","\\n").Replace("\t","\\t")+"\"";}
        if(o is bool) return ((bool)o)?"true":"false";
        if(o is Dictionary<string,object>) {
            var d=(Dictionary<string,object>)o;
            var pad=new string(' ',(indent+1)*2); var pad0=new string(' ',indent*2);
            var items=d.Select(kv=>pad+Encode(kv.Key,0)+": "+Encode(kv.Value,indent+1));
            return "{\n"+string.Join(",\n",items)+"\n"+pad0+"}";
        }
        if(o is List<object>) {
            var a=(List<object>)o;
            if(a.All(x=>x is string)&&a.Count<=8) return "["+string.Join(", ",a.Select(x=>Encode(x,0)))+"]";
            var pad=new string(' ',(indent+1)*2); var pad0=new string(' ',indent*2);
            return "[\n"+string.Join(",\n",a.Select(x=>pad+Encode(x,indent+1)))+"\n"+pad0+"]";
        }
        return o.ToString();
    }
}

// ══════════ Lab Parsing ══════════
static class Lab {
    public class Item { public string Key,Pattern,Disp; public bool On; }
    public static List<Item> Items = new List<Item> {
        // ── Order: AC, HbA1c, BUN/Cr, Na/K, ALT/AST, Alb, TC/TG/LD/HD, UA, Ca/IP, Hb, UAC, UPC ──
        new Item{Key="GluAC",Pattern=@"\bGlu",           Disp="AC",   On=true},
        new Item{Key="HbA1c",Pattern=@"\bHbA1[Cc]\b",      Disp="HbA1c",On=true},
        new Item{Key="BUN",  Pattern=@"\bBUN\b",         Disp="BUN",  On=true},
        new Item{Key="Cr",   Pattern=@"\bCr\b(?!P|E)",   Disp="Cr",   On=true},
        new Item{Key="eGFR", Pattern=@"eGFR",            Disp="eGFR", On=false},
        new Item{Key="Na",   Pattern=@"\bNa\b(?!m)",     Disp="Na",   On=true},
        new Item{Key="K",    Pattern=@"(?<![A-Za-z])K(?=\s)",Disp="K",On=true},
        new Item{Key="CRP",  Pattern=@"\bCRP\b",         Disp="CRP",  On=false},
        new Item{Key="ALT",  Pattern=@"\bALT\b",         Disp="ALT",  On=true},
        new Item{Key="AST",  Pattern=@"\bAST\b",         Disp="AST",  On=true},
        new Item{Key="Alb",  Pattern=@"\bAlbumin\b(?!.*Cr)",Disp="Alb",On=true},
        new Item{Key="Chol", Pattern=@"\bCholesterol\b", Disp="TC",   On=true},
        new Item{Key="TG",   Pattern=@"\bTG\b",          Disp="TG",   On=true},
        new Item{Key="LDL",  Pattern=@"\bLDL",           Disp="LD",   On=true},
        new Item{Key="HDL",  Pattern=@"\bHDL",           Disp="HD",   On=true},
        new Item{Key="UA",   Pattern=@"\bUric\b",        Disp="UA",   On=true},
        new Item{Key="Ca",   Pattern=@"(?<![a-z])Ca\b",  Disp="Ca",   On=true},
        new Item{Key="P",    Pattern=@"\bPhosphate\b|(?<=\s)P(?=\s{2,})",Disp="IP",On=true},
        new Item{Key="WBC",  Pattern=@"\bWBC\b",         Disp="WBC",  On=true},
        new Item{Key="PLT",  Pattern=@"\bPlatelet\b",    Disp="PLT",  On=true},
        new Item{Key="TBil", Pattern=@"\bT-Bil\b",       Disp="TBil", On=true},
        new Item{Key="Mg",   Pattern=@"\bMg\b",          Disp="Mg",   On=true},
        new Item{Key="Hb",   Pattern=@"\bHb\b(?!A)",     Disp="Hb",   On=true},
        new Item{Key="UAC",  Pattern=@"\bUACR\b|\bACR\b",   Disp="UAC",  On=true},
        new Item{Key="UPC",  Pattern=@"\bUPCR\b|\bPCR\b",   Disp="UPC",  On=true},
    };
    static string[][] Groups = {
        new[]{"BUN","Cr"}, new[]{"Na","K"}, new[]{"ALT","AST"},
        new[]{"Chol","TG","LDL","HDL"}, new[]{"Ca","P"}
    };

    public static string ExtractVal(string line, string pattern) {
        var m = Regex.Match(line, pattern);
        if (!m.Success) return null;
        var rest = line.Substring(m.Index + m.Length);
        // Match H/HH/L/LL flag with or without space before number
        var v = Regex.Match(rest, @"(?:^|\s)[HL]{1,2}\s*([\d]+\.?\d*)");
        if (v.Success) return v.Groups[1].Value;
        // Fallback: 2+ spaces then number (normal values without flag)
        v = Regex.Match(rest, @"\s{2,}([\d]+\.?\d*)");
        return v.Success ? v.Groups[1].Value : null;
    }
    public static string ExtractDate(string text) {
        var m = Regex.Match(text, "\u63a1\u6aa2\u6642\u9593[\uff1a:]?\\s*(\\d{2,3})/(\\d{2})/(\\d{2})");
        return m.Success ? m.Groups[1].Value + m.Groups[2].Value : "";
    }
    public static bool IsLabData(string text) {
        var kws = new[]{"BUN","mg/dl","mg/dL","mEq/L","U/L","mg/L","Glu","Cholesterol",
            "LDL","HDL","ALT","AST","Hemolysis","Lipemia","Icterus","Cr ",
            "\u63a1\u6aa2\u6642\u9593","\u6aa2\u9a57\u79d1","\u7d50\u679c\u503c",
            "\u53c3\u8003\u503c","\u6aa2\u9ad4","\u5831\u544a\u8005"};
        return kws.Count(k => text.Contains(k)) >= 3;
    }
    public static string Convert(string raw, HashSet<string> enabled) {
        var date = ExtractDate(raw);
        var vals = new Dictionary<string,string>();
        // Separate blood vs urine: track current section
        // UAC/UPC come from urine, everything else from blood
        var urineKeys = new HashSet<string>{"UAC","UPC"};
        bool inUrine = false;
        foreach (var line in raw.Split('\n')) {
            // Detect section changes
            if (line.Contains("\u6aa2\u9ad4")) { // 檢體
                inUrine = line.Contains("\u5c3f\u6db2"); // 尿液
            }
            foreach (var it in Items) {
                if (vals.ContainsKey(it.Key)) continue;
                // Skip: urine items only from urine section, blood items only from blood section
                bool isUrineItem = urineKeys.Contains(it.Key);
                if (isUrineItem && !inUrine) continue;
                if (!isUrineItem && inUrine) continue;
                var v = ExtractVal(line, it.Pattern);
                if (v != null) vals[it.Key] = v;
            }
        }
        // BUN: round to integer
        if (vals.ContainsKey("BUN")) {
            double bun; if(double.TryParse(vals["BUN"], out bun)) vals["BUN"]=((int)Math.Round(bun)).ToString();
        }
        var parts = new List<string>();
        var used = new HashSet<string>();
        foreach (var it in Items) {
            if (used.Contains(it.Key)||!enabled.Contains(it.Key)||!vals.ContainsKey(it.Key)) continue;
            // Special: AC with HbA1c in parentheses
            if (it.Key=="GluAC") {
                var ac=vals["GluAC"];
                if (enabled.Contains("HbA1c")&&vals.ContainsKey("HbA1c")) {
                    parts.Add("AC:"+ac+"("+vals["HbA1c"]+")");
                    used.Add("GluAC"); used.Add("HbA1c");
                } else {
                    parts.Add("AC:"+ac); used.Add("GluAC");
                }
                continue;
            }
            // HbA1c alone (only if AC didn't already consume it)
            if (it.Key=="HbA1c") {
                if (!used.Contains("HbA1c")) { parts.Add("HbA1c:"+vals["HbA1c"]); used.Add("HbA1c"); }
                continue;
            }
            bool grouped = false;
            foreach (var g in Groups) {
                if (it.Key == g[0]) {
                    var active = g.Where(k=>enabled.Contains(k)&&vals.ContainsKey(k)).ToArray();
                    if (active.Length > 1) {
                        var label = string.Join("/", active.Select(k=>Items.First(x=>x.Key==k).Disp));
                        var v = string.Join("/", active.Select(k=>vals[k]));
                        parts.Add(label+":"+v); foreach(var k in active)used.Add(k); grouped=true;
                    } break;
                } else if (g.Contains(it.Key)) {
                    if (enabled.Contains(g[0])&&vals.ContainsKey(g[0])) grouped=true; break;
                }
            }
            if (!grouped&&!used.Contains(it.Key)){parts.Add(it.Disp+":"+vals[it.Key]);used.Add(it.Key);}
        }
        if (parts.Count==0) return null;
        var r = string.Join(",", parts);
        return date!="" ? date+" "+r : r;
    }
}

// ══════════ Config ══════════
class Slot {
    public string Name="",Type="none",Text="",Hotkey="";
    public List<string> LabItems=new List<string>();
}
class Config {
    public List<Slot> Slots = new List<Slot>();
    static string _path;
    static string _dir;
    static string ConfigDir {
        get {
            if(_dir==null){
                _dir=System.IO.Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                    "EMR_Report");
                if(!Directory.Exists(_dir)) Directory.CreateDirectory(_dir);
            }
            return _dir;
        }
    }
    static string Path {
        get {
            if(_path==null) _path=System.IO.Path.Combine(ConfigDir,"lab_formatter_config.json");
            return _path;
        }
    }

    public void Load() {
        Slots.Clear();
        try {
            if (!File.Exists(Path)) { CreateDefault(); }
            var d = Json.Parse(File.ReadAllText(Path, System.Text.Encoding.UTF8));
            if (d.ContainsKey("slots") && d["slots"] is List<object>) {
                var arr=(List<object>)d["slots"];
                foreach (var obj in arr) {
                    var s=(Dictionary<string,object>)obj;
                    var slot = new Slot();
                    if(s.ContainsKey("name")) slot.Name=s["name"].ToString();
                    if(s.ContainsKey("type")) slot.Type=s["type"].ToString();
                    if(s.ContainsKey("text")&&s["text"]!=null) slot.Text=s["text"].ToString();
                    if(s.ContainsKey("_hotkey")) slot.Hotkey=s["_hotkey"].ToString();
                    if(s.ContainsKey("lab_items")&&s["lab_items"] is List<object>){
                        var li=(List<object>)s["lab_items"];
                        slot.LabItems=li.Select(x=>x.ToString()).ToList();}
                    Slots.Add(slot);
                }
            }
        } catch { }
        while(Slots.Count<4) {
            var i=Slots.Count;
            if(i==2) Slots.Add(new Slot{Name="\u6fc3\u7e2e\u5831\u544a",Type="lab",
                Hotkey="Ctrl+"+(i+1),
                LabItems=Lab.Items.Where(x=>x.On).Select(x=>x.Key).ToList()});
            else Slots.Add(new Slot{Name="\u5feb\u8cbc"+(i<2?i+1:i),Type="paste",
                Hotkey="Ctrl+"+(i+1)});
        }
    }
    public void Save() {
        var doc = new Dictionary<string,object>();
        doc["\u005f\u8aaa\u660e"]="Lab Data Formatter \u8a2d\u5b9a\u6a94";
        var fmt=new Dictionary<string,object>();
        fmt["type"]=new List<object>{"paste=\u5feb\u8cbc\u6587\u5b57","lab=\u5831\u544a\u6574\u7406",
            "template=\u7bc4\u672c\u5957\u7528","none=\u672a\u555f\u7528"};
        fmt["\u71b1\u9375"]="Ctrl+0=\u8a2d\u5b9a, Ctrl+1~4=slots[0]~[3]";
        doc["\u005f\u683c\u5f0f\u8aaa\u660e"]=fmt;
        var arr = new List<object>();
        for(int i=0;i<Slots.Count;i++){var s=Slots[i];
            var d=new Dictionary<string,object>();
            d["_hotkey"]="Ctrl+"+(i+1); d["name"]=s.Name; d["type"]=s.Type;
            if(s.Type=="paste"||s.Type=="template") d["text"]=s.Text;
            if(s.Type=="lab") d["lab_items"]=s.LabItems.Cast<object>().ToList();
            arr.Add(d);}
        doc["slots"]=arr;
        File.WriteAllText(Path, Json.Encode(doc,0), System.Text.Encoding.UTF8);
    }
    void CreateDefault() {
        var defaults = new List<Slot>();
        defaults.Add(new Slot{Name="\u5feb\u8cbc1",Type="paste",Hotkey="Ctrl+1"});
        defaults.Add(new Slot{Name="\u5feb\u8cbc2",Type="paste",Hotkey="Ctrl+2"});
        defaults.Add(new Slot{Name="\u6fc3\u7e2e\u5831\u544a",Type="lab",Hotkey="Ctrl+3",
            LabItems=Lab.Items.Where(x=>x.On).Select(x=>x.Key).ToList()});
        defaults.Add(new Slot{Name="\u5feb\u8cbc3",Type="paste",Hotkey="Ctrl+4"});
        var doc = new Dictionary<string,object>();
        doc["\u005f\u8aaa\u660e"]="Lab Data Formatter \u8a2d\u5b9a\u6a94";
        var fmt=new Dictionary<string,object>();
        fmt["type"]=new List<object>{"paste=\u5feb\u8cbc\u6587\u5b57","lab=\u5831\u544a\u6574\u7406",
            "template=\u7bc4\u672c\u5957\u7528","none=\u672a\u555f\u7528"};
        fmt["\u71b1\u9375"]="Ctrl+0=\u8a2d\u5b9a, Ctrl+1~4=slots[0]~[3]";
        doc["\u005f\u683c\u5f0f\u8aaa\u660e"]=fmt;
        var arr = new List<object>();
        for(int i=0;i<defaults.Count;i++){var s=defaults[i];
            var d=new Dictionary<string,object>();
            d["_hotkey"]="Ctrl+"+(i+1); d["name"]=s.Name; d["type"]=s.Type;
            if(s.Type=="paste"||s.Type=="template") d["text"]=s.Text;
            if(s.Type=="lab") d["lab_items"]=s.LabItems.Cast<object>().ToList();
            arr.Add(d);}
        doc["slots"]=arr;
        try{File.WriteAllText(Path, Json.Encode(doc,0), System.Text.Encoding.UTF8);}catch{}
    }
    public HashSet<string> GetLabItems() {
        var s=Slots.FirstOrDefault(x=>x.Type=="lab");
        if(s!=null&&s.LabItems.Count>0) return new HashSet<string>(s.LabItems);
        return new HashSet<string>(Lab.Items.Where(x=>x.On).Select(x=>x.Key));
    }
    public static void OpenJson() { try{System.Diagnostics.Process.Start(Path);}catch{} }
    public static void OpenFolder() { try{System.Diagnostics.Process.Start("explorer.exe",
        ConfigDir);}catch{} }
    public static string JsonPath { get { return Path; } }
}

// ══════════ Native helpers ══════════
static class NativeMethods {
    [DllImport("user32")] public static extern bool ReleaseCapture();
    [DllImport("user32")] public static extern int SendMessage(IntPtr h, int msg, int wp, int lp);
}

// ══════════ Main App ══════════
class App : Form {
    const string VER="v1.1.1";
    [DllImport("user32")] static extern bool RegisterHotKey(IntPtr h,int id,uint mod,uint vk);
    [DllImport("user32")] static extern bool UnregisterHotKey(IntPtr h,int id);
    [DllImport("user32")] static extern void keybd_event(byte vk,byte scan,uint flags,UIntPtr extra);
    [DllImport("user32")] static extern bool DestroyIcon(IntPtr handle);

    NotifyIcon tray; Config cfg = new Config();
    string labResult; string lastClip; DateTime ignoreUntil;
    System.Windows.Forms.Timer clipTimer;
    const uint MOD_CTRL=2; const int WM_HOTKEY=0x312;
    static readonly Dictionary<int,uint> HK = new Dictionary<int,uint>{
        {0,0x30},{1,0x31},{2,0x32},{3,0x33},{4,0x34}};

    // ── Multi-report merge buffer (10s window) ──
    List<string> rawBuffer = new List<string>();  // stores raw clipboard text
    System.Windows.Forms.Timer mergeTimer;

    Icon MakeIcon(Color c) {
        var bmp=new Bitmap(32,32); using(var g=Graphics.FromImage(bmp)){
            g.SmoothingMode=System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            g.Clear(Color.Transparent);
            using(var b=new SolidBrush(Color.FromArgb(60,c)))g.FillEllipse(b,2,0,28,28);
            using(var b=new SolidBrush(c))g.FillEllipse(b,7,3,18,18);
            using(var b=new SolidBrush(ControlPaint.Dark(c)))g.FillRectangle(b,11,20,10,5);
            using(var b=new SolidBrush(ControlPaint.DarkDark(c)))g.FillRectangle(b,12,25,8,4);
            using(var p=new Pen(Color.FromArgb(180,255,255,255),1.5f)){
                g.DrawLine(p,14,14,15,7); g.DrawLine(p,15,7,17,12);
                g.DrawLine(p,17,12,19,7); g.DrawLine(p,19,7,20,14);}
        }
        var hIcon=bmp.GetHicon(); var icon=Icon.FromHandle(hIcon);
        var clone=(Icon)icon.Clone(); DestroyIcon(hIcon); bmp.Dispose();
        return clone;
    }

    public App() {
        cfg.Load(); ShowInTaskbar=false; WindowState=FormWindowState.Minimized; Visible=false;
        tray=new NotifyIcon{Icon=MakeIcon(Color.Gold),Text="\u7b49\u5f85 Ctrl+C...",Visible=true};
        var menu=new ContextMenuStrip();
        menu.Items.Add("Ctrl+0 \u8a2d\u5b9a",null,(s,e)=>ShowSettings());
        menu.Items.Add("\u7de8\u8f2f JSON",null,(s,e)=>Config.OpenJson());
        menu.Items.Add("\u624b\u52d5\u8f49\u63db",null,(s,e)=>ShowManual());
        menu.Items.Add("\u4f7f\u7528\u8aaa\u660e",null,(s,e)=>OpenReadme());
        menu.Items.Add(new ToolStripSeparator());
        menu.Items.Add("\u7d50\u675f",null,(s,e)=>{tray.Visible=false;Application.Exit();});
        tray.ContextMenuStrip=menu;
        tray.DoubleClick+=(s,e)=>ShowSettings();

        foreach(var kv in HK) RegisterHotKey(Handle,kv.Key,MOD_CTRL,kv.Value);

        lastClip=TryGetClip(); ignoreUntil=DateTime.MinValue;
        clipTimer=new System.Windows.Forms.Timer{Interval=400};
        clipTimer.Tick+=(s,e)=>PollClip();
        clipTimer.Start();
    }

    string TryGetClip(){try{return Clipboard.ContainsText()?Clipboard.GetText():"";}catch{return"";}}

    void PollClip() {
        var t=TryGetClip();
        if(t!=""&&t!=lastClip){lastClip=t;
            if(DateTime.Now<ignoreUntil)return;
            if(!Lab.IsLabData(t))return;
            // Add raw text to buffer
            rawBuffer.Add(t);
            // Convert ALL buffered raw text as one combined report
            var combined=string.Join("\n",rawBuffer);
            var r=Lab.Convert(combined,cfg.GetLabItems());
            if(r!=null){
                labResult=r;
                var countTag=rawBuffer.Count>1?
                    " ["+rawBuffer.Count+"\u7b46]":"";
                SetTray(Color.LimeGreen,"\u2713"+countTag+" "+r.Substring(0,Math.Min(50,r.Length)));
                ShowPreview(r, rawBuffer.Count);
                ResetMergeTimer();
            }}
    }
    void ResetMergeTimer() {
        if(mergeTimer==null){
            mergeTimer=new System.Windows.Forms.Timer{Interval=10000};
            mergeTimer.Tick+=(s,e)=>{mergeTimer.Stop(); FinalizeMerge();};}
        mergeTimer.Stop(); mergeTimer.Start();
    }
    void FinalizeMerge() {
        rawBuffer.Clear();
        DelayYellow();
    }

    // ══════ Floating Preview Popup ══════
    Form previewForm;
    Label previewLabel;
    Label previewTitle;
    System.Windows.Forms.Timer previewTimer;
    void ShowPreview(string result, int count) {
        if(previewForm==null||previewForm.IsDisposed) {
            previewForm=new Form{
                FormBorderStyle=FormBorderStyle.None,
                ShowInTaskbar=false,
                TopMost=true,
                BackColor=Color.FromArgb(45,45,45),
                Opacity=0.95,
                StartPosition=FormStartPosition.Manual,
                Size=new Size(500,70)
            };
            var topBar=new Panel{Dock=DockStyle.Top,Height=22,BackColor=Color.FromArgb(76,175,80)};
            previewTitle=new Label{
                ForeColor=Color.White,Font=new Font("Microsoft JhengHei UI",9,FontStyle.Bold),
                Dock=DockStyle.Fill,TextAlign=ContentAlignment.MiddleLeft,Padding=new Padding(8,0,0,0)};
            var closeBtn=new Label{Text="\u2715",ForeColor=Color.White,
                Font=new Font("Arial",9),Dock=DockStyle.Right,Width=28,
                TextAlign=ContentAlignment.MiddleCenter,Cursor=Cursors.Hand};
            closeBtn.Click+=(s,e)=>previewForm.Hide();
            topBar.Controls.Add(previewTitle); topBar.Controls.Add(closeBtn);
            previewLabel=new Label{ForeColor=Color.FromArgb(200,255,200),
                Font=new Font("Consolas",10),Dock=DockStyle.Fill,
                TextAlign=ContentAlignment.TopLeft,Padding=new Padding(10,6,10,6)};
            previewForm.Controls.Add(previewLabel);
            previewForm.Controls.Add(topBar);
            topBar.MouseDown+=(s,e)=>{if(e.Button==MouseButtons.Left){
                NativeMethods.ReleaseCapture();NativeMethods.SendMessage(previewForm.Handle,0xA1,0x2,0);}};
            previewTitle.MouseDown+=(s,e)=>{if(e.Button==MouseButtons.Left){
                NativeMethods.ReleaseCapture();NativeMethods.SendMessage(previewForm.Handle,0xA1,0x2,0);}};
        }
        // Update title with count
        if(count>1)
            previewTitle.Text="\u2713 "+count+"\u7b46\u5831\u544a\u5df2\u8f49\u63db\uff0c10\u79d2\u5167\u53ef\u7e7c\u7e8c\u8907\u88fd | Ctrl+3 \u8cbc\u4e0a";
        else
            previewTitle.Text="\u2713 \u5831\u544a\u5df2\u8f49\u63db\uff0c\u6309 Ctrl+3 \u8cbc\u4e0a";

        previewLabel.Text=result;
        // Size: width by longest line, height by line count
        var lines=result.Split('\n');
        int maxLen=0; foreach(var ln in lines) if(ln.Length>maxLen) maxLen=ln.Length;
        int textW=(int)(maxLen*7.5)+40;
        if(textW<400) textW=400;
        if(textW>Screen.PrimaryScreen.WorkingArea.Width-20) textW=Screen.PrimaryScreen.WorkingArea.Width-20;
        int textH=22+Math.Max(lines.Length,1)*20+16; // topBar + lines + padding
        if(textH<60) textH=60;
        if(textH>300) textH=300;
        previewForm.Size=new Size(textW,textH);
        // Position: bottom-right above taskbar
        var wa=Screen.PrimaryScreen.WorkingArea;
        previewForm.Location=new Point(wa.Right-previewForm.Width-10, wa.Bottom-previewForm.Height-10);
        previewForm.Show();
        // Auto-hide after 12 seconds (longer for multi-report)
        int hideDelay=count>1?15000:8000;
        if(previewTimer==null){previewTimer=new System.Windows.Forms.Timer();
            previewTimer.Tick+=(s,e)=>{previewTimer.Stop();if(previewForm!=null&&!previewForm.IsDisposed)previewForm.Hide();};}
        previewTimer.Interval=hideDelay;
        previewTimer.Stop(); previewTimer.Start();
    }
    // Hide preview when Ctrl+3 pastes
    void HidePreview(){
        if(previewForm!=null&&!previewForm.IsDisposed&&previewForm.Visible) previewForm.Hide();
        if(previewTimer!=null) previewTimer.Stop();
    }
    void SetTray(Color c,string tip){
        var old=tray.Icon; tray.Icon=MakeIcon(c); if(old!=null)old.Dispose();
        tray.Text=tip.Length>63?tip.Substring(0,63):tip;}
    System.Windows.Forms.Timer yellowTimer;
    void DelayYellow(){
        if(yellowTimer==null){yellowTimer=new System.Windows.Forms.Timer{Interval=3000};
            yellowTimer.Tick+=(s,e)=>{yellowTimer.Stop();
                SetTray(Color.Gold,labResult!=null?
                    "Ctrl+3 \u8cbc\u4e0a | "+labResult.Substring(0,Math.Min(40,labResult.Length)):
                    "\u7b49\u5f85 Ctrl+C...");};}
        yellowTimer.Stop(); yellowTimer.Start();
    }
    void WriteClip(string t){ignoreUntil=DateTime.Now.AddMilliseconds(800);
        try{Clipboard.SetText(t);}catch{} lastClip=t;}
    void SimCtrlV(){Thread.Sleep(50);
        keybd_event(0x12,0,2,UIntPtr.Zero);keybd_event(0x11,0,2,UIntPtr.Zero);keybd_event(0x10,0,2,UIntPtr.Zero);
        Thread.Sleep(30);
        keybd_event(0x11,0,0,UIntPtr.Zero);keybd_event(0x56,0,0,UIntPtr.Zero);
        keybd_event(0x56,0,2,UIntPtr.Zero);keybd_event(0x11,0,2,UIntPtr.Zero);}

    protected override void WndProc(ref Message m) {
        if(m.Msg==WM_HOTKEY){int id=(int)m.WParam;
            if(id==0){ShowSettings();return;}
            int idx=id-1; if(idx<0||idx>=cfg.Slots.Count)return;
            var s=cfg.Slots[idx];
            try{
                if(s.Type=="paste"){if(s.Text!=""){WriteClip(s.Text);SimCtrlV();}}
                else if(s.Type=="lab"){
                    // Finalize any pending buffer
                    if(rawBuffer.Count>0){
                        var combined=string.Join("\n",rawBuffer);
                        labResult=Lab.Convert(combined,cfg.GetLabItems());
                        rawBuffer.Clear();
                        if(mergeTimer!=null) mergeTimer.Stop();
                    }
                    if(labResult!=null){WriteClip(labResult);SimCtrlV();
                    HidePreview();SetTray(Color.LimeGreen,"\u2713 \u5df2\u8cbc\u4e0a");DelayYellow();}
                    else{SetTray(Color.Red,"\u8acb\u5148 Ctrl+C \u8907\u88fd\u5831\u544a");DelayYellow();}}
                else if(s.Type=="template"&&s.Text!=""){
                    var clip=TryGetClip();WriteClip(s.Text.Replace("{clipboard}",clip));SimCtrlV();}
            }catch{}}
        base.WndProc(ref m);
    }
    protected override void OnFormClosed(FormClosedEventArgs e){
        foreach(var k in HK.Keys)UnregisterHotKey(Handle,k);tray.Visible=false;base.OnFormClosed(e);}
    protected override void SetVisibleCore(bool v){base.SetVisibleCore(false);}

    // ══════ Settings Window ══════
    Form settingsForm;
    void ShowSettings() {
        if(settingsForm!=null&&!settingsForm.IsDisposed){settingsForm.BringToFront();return;}
        cfg.Load();
        var f=new Form{Text="Lab Formatter "+VER+" \u8a2d\u5b9a",Size=new Size(620,660),
            StartPosition=FormStartPosition.CenterScreen,TopMost=true,
            MaximizeBox=false,Font=new Font("Microsoft JhengHei UI",9)};
        settingsForm=f;
        int fw=f.ClientSize.Width;

        var lbl=new Label{Text="Ctrl+1~2,4 \u5feb\u901f\u8cbc\u4e0a\u6587\u5b57 | Ctrl+C \u8907\u88fd\u5831\u544a \u2192 Ctrl+3 \u8cbc\u4e0a\u6fc3\u7e2e\u5831\u544a",
            Left=0,Top=0,Width=fw,Height=22,TextAlign=ContentAlignment.MiddleCenter,ForeColor=Color.Gray};
        f.Controls.Add(lbl);

        // 4 slot editors stacked vertically
        var slots=new InlineSlot[4];
        int sy=26;
        for(int i=0;i<4;i++){
            slots[i]=new InlineSlot(f,cfg.Slots[i],i,sy,fw-20);
            sy+=slots[i].Height+4;
        }

        // Save button
        var saveBtn=new Button{Text="\u5132\u5b58\u8a2d\u5b9a",Left=fw/2-120,Top=sy+2,Width=120,Height=30,
            BackColor=Color.FromArgb(76,175,80),ForeColor=Color.White,
            Font=new Font("Microsoft JhengHei UI",10,FontStyle.Bold)};
        saveBtn.Click+=(s,e)=>{
            for(int i=0;i<4;i++) slots[i].SaveTo(cfg.Slots[i]);
            cfg.Save(); SetTray(Color.LimeGreen,"\u2713 \u5df2\u5132\u5b58");DelayYellow();
        };
        var helpBtn=new Button{Text="\u4f7f\u7528\u8aaa\u660e",Left=fw/2+10,Top=sy+2,Width=110,Height=30,
            BackColor=Color.FromArgb(33,150,243),ForeColor=Color.White,
            Font=new Font("Microsoft JhengHei UI",10)};
        helpBtn.Click+=(s,e)=>OpenReadme();
        f.Controls.Add(saveBtn);f.Controls.Add(helpBtn);

        // JSON row
        int jy=sy+38;
        var jlbl=new Label{Text="JSON:",Left=10,Top=jy+2,Width=38,ForeColor=Color.Gray};
        var pe=new TextBox{Left=50,Top=jy,Width=fw-200,ReadOnly=true,Text=Config.JsonPath,
            BackColor=Color.WhiteSmoke,Font=new Font("Consolas",8)};
        var b1=new Button{Text="\u958b\u555f",Left=fw-144,Top=jy-1,Width=50,Height=22};b1.Click+=(s,e)=>Config.OpenJson();
        var b2=new Button{Text="\u8b80\u53d6",Left=fw-90,Top=jy-1,Width=50,Height=22};
        b2.Click+=(s,e)=>{cfg.Load();for(int i=0;i<4;i++)slots[i].LoadFrom(cfg.Slots[i]);
            SetTray(Color.LimeGreen,"\u2713 \u5df2\u8b80\u53d6");DelayYellow();};
        f.Controls.AddRange(new Control[]{jlbl,pe,b1,b2});

        // Result
        int ry=jy+26;
        f.Controls.Add(new Label{Text="\u6700\u65b0:",Left=10,Top=ry+2,Width=40,ForeColor=Color.DarkGreen,
            Font=new Font("Microsoft JhengHei UI",9,FontStyle.Bold)});
        var st=new TextBox{Text=labResult!=null?labResult:"\u5c1a\u7121\u7d50\u679c",
            Left=52,Top=ry,Width=fw-62,Height=20,ReadOnly=true,BackColor=Color.FromArgb(240,255,240),
            ForeColor=Color.DarkGreen,Font=new Font("Consolas",9),BorderStyle=BorderStyle.FixedSingle};
        f.Controls.Add(st);

        // Version + feedback
        int fy=ry+28;
        var feedback=new Label{Text="\u610f\u898b\u56de\u994b: \u5433\u5cb3\u9716\u91ab\u5e2b  DAL93@tpech.gov.tw    "+VER,
            Left=0,Top=fy,Width=fw,Height=18,TextAlign=ContentAlignment.MiddleRight,
            ForeColor=Color.FromArgb(160,160,160),Font=new Font("Microsoft JhengHei UI",8),
            Padding=new Padding(0,0,10,0)};
        f.Controls.Add(feedback);

        f.ClientSize=new Size(fw,fy+20);
        f.ShowDialog();
        settingsForm=null;
    }

    // ── Inline slot editor (one row per slot) ──
    class InlineSlot {
        ComboBox combo; TextBox nameBox, txtBox;
        string[] types={"none","paste","lab","template"};
        string[] tnames={"\u672a\u8a2d\u5b9a","\u5feb\u8cbc","\u5831\u544a","\u7bc4\u672c"};
        GroupBox gb;
        public int Height { get { return gb.Height; } }

        public InlineSlot(Form f, Slot s, int idx, int top, int width) {
            bool isLab=(s.Type=="lab");
            int h=isLab?48:80;
            gb=new GroupBox{Text="Ctrl+"+(idx+1),Left=10,Top=top,Width=width,Height=h,
                Font=new Font("Microsoft JhengHei UI",9)};
            f.Controls.Add(gb);

            // Row 1: name + type
            nameBox=new TextBox{Left=50,Top=16,Width=150,Text=s.Name,Font=new Font("Microsoft JhengHei UI",9)};
            gb.Controls.Add(new Label{Text="\u540d\u7a31:",Left=8,Top=19,Width=40,AutoSize=false});
            gb.Controls.Add(nameBox);
            combo=new ComboBox{Left=260,Top=16,Width=80,DropDownStyle=ComboBoxStyle.DropDownList,
                Font=new Font("Microsoft JhengHei UI",9)};
            combo.Items.AddRange(tnames);combo.SelectedIndex=Array.IndexOf(types,s.Type);
            gb.Controls.Add(new Label{Text="\u985e\u578b:",Left=218,Top=19,Width=40,AutoSize=false});
            gb.Controls.Add(combo);

            // Row 2: content (only for paste/template)
            if(!isLab){
                txtBox=new TextBox{Left=50,Top=42,Width=width-70,Height=30,
                    Multiline=true,ScrollBars=ScrollBars.Vertical,
                    Text=s.Text,Font=new Font("Consolas",9)};
                gb.Controls.Add(new Label{Text="\u5167\u5bb9:",Left=8,Top=45,Width=40,AutoSize=false});
                gb.Controls.Add(txtBox);
            } else {
                gb.Controls.Add(new Label{Text="Ctrl+C \u8907\u88fd\u5831\u544a \u2192 Ctrl+"+(idx+1)+" \u8cbc\u4e0a\u6fc3\u7e2e\u5831\u544a",
                    Left=360,Top=19,Width=240,ForeColor=Color.Gray,AutoSize=false});
            }

            combo.SelectedIndexChanged+=(x,y)=>{
                int ti=combo.SelectedIndex;
                bool nowLab=(ti==2);
                if(nowLab&&txtBox!=null){txtBox.Visible=false;gb.Height=48;}
                else if(!nowLab){
                    if(txtBox==null){
                        txtBox=new TextBox{Left=50,Top=42,Width=width-70,Height=30,
                            Multiline=true,ScrollBars=ScrollBars.Vertical,Font=new Font("Consolas",9)};
                        gb.Controls.Add(new Label{Text="\u5167\u5bb9:",Left=8,Top=45,Width=40,AutoSize=false});
                        gb.Controls.Add(txtBox);
                    }
                    txtBox.Visible=true;gb.Height=80;
                }
            };
        }
        public void SaveTo(Slot s) {
            s.Name=nameBox.Text.Trim(); if(s.Name=="")s.Name="Slot";
            s.Type=types[combo.SelectedIndex];
            if(s.Type=="paste"||s.Type=="template"){s.Text=txtBox!=null?txtBox.Text:"";s.LabItems.Clear();}
            else if(s.Type=="lab"){s.Text="";} // keep existing lab_items
            else{s.Text="";s.LabItems.Clear();}
        }
        public void LoadFrom(Slot s) {
            nameBox.Text=s.Name;
            combo.SelectedIndex=Array.IndexOf(types,s.Type);
            if(txtBox!=null) txtBox.Text=s.Text;
        }
    }

    // ══════ Manual Window ══════
    Form manualForm;
    static void OpenReadme() {
        var dir=AppDomain.CurrentDomain.BaseDirectory;
        var path=System.IO.Path.Combine(dir,"readme.txt");
        if(!File.Exists(path)){
            // Create embedded readme if file doesn't exist
            try{File.WriteAllText(path,
                "Lab Data Formatter "+VER+"\n\n"+
                "\u8acb\u5f9e\u8a2d\u5b9a\u756b\u9762\u6216\u53f3\u9375\u9078\u55ae\u958b\u555f\u8aaa\u660e\u6a94\u3002\n"+
                "\u5982\u679c\u770b\u5230\u6b64\u8a0a\u606f\uff0c\u8acb\u5c07 readme.txt \u653e\u5728 exe \u540c\u8cc7\u6599\u593e\u3002",
                System.Text.Encoding.UTF8);}catch{}
        }
        try{System.Diagnostics.Process.Start(path);}catch{}
    }

    void ShowManual() {
        if(manualForm!=null&&!manualForm.IsDisposed){manualForm.BringToFront();return;}
        var f=new Form{Text="\u624b\u52d5\u8f49\u63db",Size=new Size(560,420),StartPosition=FormStartPosition.CenterScreen,
            TopMost=true,Font=new Font("Microsoft JhengHei UI",9)};
        manualForm=f;
        f.Controls.Add(new Label{Text="\u8cbc\u4e0a\u6aa2\u9a57\u5831\u544a\uff1a",Left=10,Top=8,Width=200});
        var inp=new TextBox{Left=10,Top=28,Width=525,Height=200,Multiline=true,ScrollBars=ScrollBars.Vertical,
            Font=new Font("Consolas",10)};f.Controls.Add(inp);
        var outBox=new TextBox{Left=10,Top=270,Width=525,Height=50,ReadOnly=true,BackColor=Color.WhiteSmoke,
            Font=new Font("Consolas",11)};f.Controls.Add(outBox);
        var b1=new Button{Text="\u8f49\u63db",Left=160,Top=235,Width=80,BackColor=Color.FromArgb(76,175,80),ForeColor=Color.White};
        b1.Click+=(s,e)=>{var r=Lab.Convert(inp.Text,cfg.GetLabItems());outBox.Text=r!=null?r:"\u672a\u627e\u5230\u6578\u503c";
            if(r!=null){labResult=r;SetTray(Color.LimeGreen,"\u2713 "+r.Substring(0,Math.Min(38,r.Length)));DelayYellow();}};
        var b2=new Button{Text="\u8907\u88fd\u7d50\u679c",Left=250,Top=235,Width=80,BackColor=Color.FromArgb(33,150,243),ForeColor=Color.White};
        b2.Click+=(s,e)=>{if(outBox.Text!="")WriteClip(outBox.Text);};
        f.Controls.AddRange(new Control[]{b1,b2});f.Show();
    }

    [STAThread] static void Main() {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Application.Run(new App());
    }
}
