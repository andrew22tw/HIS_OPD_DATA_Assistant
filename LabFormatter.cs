// Lab Data Formatter v1.4.1
// Author: \u5433\u5cb3\u9716\u91ab\u5e2b (DAL93@tpech.gov.tw)
// Compile: build.bat (auto-finds csc.exe)
// Hotkeys: Alt+1=Capture, Alt+2=Paste, Ctrl+0=Settings, Ctrl+1~4=Custom slots
// Architecture: Event-driven clipboard (WM_CLIPBOARDUPDATE), SendInput, window-aware capture

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
    public class Item {
        public string Key,Pattern,Disp; public bool On;
        public Regex Rx; // Precompiled regex
    }
    // Precompiled value extraction patterns
    static readonly Regex RxFlag = new Regex(@"(?:^|\s)[HL]{1,2}\s*([\d]+\.?\d*)", RegexOptions.Compiled);
    static readonly Regex RxVal = new Regex(@"\s{2,}([\d]+\.?\d*)", RegexOptions.Compiled);

    public static List<Item> Items = new List<Item> {
        // ── Order: AC,HbA1c,BUN/Cr,Na/K,HCO3,ALT/AST,Alb,TC/TG/L/H,UA,Ca/IP,Iron/TIBC/Ferritin,Hb,ACR,PCR ──
        new Item{Key="GluAC",Pattern=@"\bGlu",           Disp="AC",   On=true},
        new Item{Key="HbA1c",Pattern=@"\bHbA1[Cc]\b",    Disp="HbA1c",On=true},
        new Item{Key="BUN",  Pattern=@"\bBUN\b",         Disp="BUN",  On=true},
        new Item{Key="Cr",   Pattern=@"\bCr\b(?!P|E)",   Disp="Cr",   On=true},
        new Item{Key="eGFR", Pattern=@"eGFR",            Disp="eGFR", On=false},
        new Item{Key="Na",   Pattern=@"\bNa\b(?!m)",     Disp="Na",   On=true},
        new Item{Key="K",    Pattern=@"(?<![A-Za-z])K(?=\s)",Disp="K",On=true},
        new Item{Key="HCO3", Pattern=@"\bHCO3\b",        Disp="HCO3", On=true},
        new Item{Key="CRP",  Pattern=@"\bCRP\b",         Disp="CRP",  On=false},
        new Item{Key="ALT",  Pattern=@"\bALT\b",         Disp="ALT",  On=true},
        new Item{Key="AST",  Pattern=@"\bAST\b",         Disp="AST",  On=true},
        new Item{Key="Alb",  Pattern=@"\bAlbumin\b(?!.*Cr)",Disp="Alb",On=true},
        new Item{Key="Chol", Pattern=@"\bCholesterol\b", Disp="TC",   On=true},
        new Item{Key="TG",   Pattern=@"\bTG\b",          Disp="TG",   On=true},
        new Item{Key="LDL",  Pattern=@"\bLDL",           Disp="L",    On=true},
        new Item{Key="HDL",  Pattern=@"\bHDL",           Disp="H",    On=true},
        new Item{Key="UA",   Pattern=@"\bUric\b",        Disp="UA",   On=true},
        new Item{Key="Ca",   Pattern=@"(?<![a-z])Ca\b",  Disp="Ca",   On=true},
        new Item{Key="P",    Pattern=@"\bPhosphate\b|(?<=\s)P(?=\s{2,})",Disp="IP",On=true},
        new Item{Key="Iron", Pattern=@"\bIron\b",        Disp="Iron", On=true},
        new Item{Key="TIBC", Pattern=@"\bTIBC\b",        Disp="TIBC", On=true},
        new Item{Key="Ferritin",Pattern=@"\bFerritin\b", Disp="Ferritin",On=true},
        new Item{Key="WBC",  Pattern=@"\bWBC\b",         Disp="WBC",  On=true},
        new Item{Key="PLT",  Pattern=@"\bPlatelet\b",    Disp="PLT",  On=true},
        new Item{Key="TBil", Pattern=@"\bT-Bil\b",       Disp="TBil", On=true},
        new Item{Key="Mg",   Pattern=@"\bMg\b",          Disp="Mg",   On=true},
        new Item{Key="Hb",   Pattern=@"\bHb\b(?!A)",     Disp="Hb",   On=true},
        new Item{Key="UAC",  Pattern=@"\bUACR\b|\bACR\b|\bA/C\b",Disp="ACR",On=true},
        new Item{Key="UPC",  Pattern=@"\bUPCR\b|\bPCR\b|\bP/C\b",Disp="PCR",On=true},
    };
    // Precompile all item regexes at startup
    static Lab() { foreach(var it in Items) it.Rx=new Regex(it.Pattern, RegexOptions.Compiled); }
    static string[][] Groups = {
        new[]{"BUN","Cr"}, new[]{"Na","K"}, new[]{"ALT","AST"},
        new[]{"Chol","TG","LDL","HDL"}, new[]{"Ca","P"},
        new[]{"Iron","TIBC","Ferritin"}
    };

    public static string ExtractVal(string line, Regex rx) {
        var m = rx.Match(line);
        if (!m.Success) return null;
        var rest = line.Substring(m.Index + m.Length);
        var v = RxFlag.Match(rest);
        if (v.Success) return v.Groups[1].Value;
        v = RxVal.Match(rest);
        return v.Success ? v.Groups[1].Value : null;
    }
    static readonly Regex RxDate = new Regex("\u63a1\u6aa2\u6642\u9593[\uff1a:]?\\s*(\\d{2,3})/(\\d{2})/(\\d{2})", RegexOptions.Compiled);
    public static string ExtractDate(string text) {
        var m = RxDate.Match(text);
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
                var v = ExtractVal(line, it.Rx);
                if (v != null) vals[it.Key] = v;
            }
        }
        // Round to integer: BUN, Iron, TIBC, Ferritin
        var roundKeys = new[]{"BUN","HCO3","Iron","TIBC","Ferritin"};
        foreach(var rk in roundKeys) {
            if(vals.ContainsKey(rk)){double d;if(double.TryParse(vals[rk],out d))vals[rk]=((int)Math.Round(d)).ToString();}
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
        if (date!="") r = date+" "+r;
        // Word-wrap at 80 chars, break at comma, indent continuation with 6 spaces
        if (r.Length>80) {
            var sb = new System.Text.StringBuilder();
            int col=0;
            var tokens=r.Split(',');
            for(int ti=0;ti<tokens.Length;ti++){
                var tok=tokens[ti]+(ti<tokens.Length-1?",":"");
                if(col==0){sb.Append(tok);col=tok.Length;}
                else if(col+tok.Length>80){
                    sb.Append("\n      ");col=6;sb.Append(tok);col+=tok.Length;}
                else{sb.Append(tok);col+=tok.Length;}
            }
            r=sb.ToString();
        }
        return r;
    }
}

// ══════════ Config ══════════
class Slot {
    public string Name="",Type="none",Text="",Hotkey="";
    public List<string> LabItems=new List<string>();
}
static class VK {
    // Modifier name -> RegisterHotKey flag
    public static uint ModFlag(string mod) { return mod=="Alt"?(uint)1:(uint)2; } // Alt=1,Ctrl=2
    // Key name -> virtual key code
    public static uint Code(string key) {
        if(key.Length==1){
            char c=char.ToUpper(key[0]);
            if(c>='A'&&c<='Z') return (uint)c;
            if(c>='0'&&c<='9') return (uint)c;
        }
        if(key=="`"||key=="~") return 0xC0;
        if(key=="-") return 0xBD; if(key=="=") return 0xBB;
        if(key=="["||key=="{") return 0xDB; if(key=="]"||key=="}") return 0xDD;
        if(key=="\\"||key=="|") return 0xDC;
        if(key==";"||key==":") return 0xBA; if(key=="'"||key=="\"") return 0xDE;
        if(key==","||key=="<") return 0xBC; if(key=="."||key==">") return 0xBE;
        if(key=="/"||key=="?") return 0xBF;
        if(key.StartsWith("F")&&key.Length>=2){int n;if(int.TryParse(key.Substring(1),out n)&&n>=1&&n<=12)return(uint)(0x6F+n);}
        return 0xC0; // default backtick
    }
    // Virtual key code -> display name
    public static string Name(uint vk) {
        if(vk>=0x41&&vk<=0x5A) return((char)vk).ToString();
        if(vk>=0x30&&vk<=0x39) return((char)vk).ToString();
        if(vk==0xC0) return "`"; if(vk==0xBD) return "-"; if(vk==0xBB) return "=";
        if(vk==0xDB) return "["; if(vk==0xDD) return "]"; if(vk==0xDC) return "\\";
        if(vk==0xBA) return ";"; if(vk==0xDE) return "'";
        if(vk==0xBC) return ","; if(vk==0xBE) return "."; if(vk==0xBF) return "/";
        if(vk>=0x70&&vk<=0x7B) return "F"+(vk-0x6F);
        return "?";
    }
}
class Config {
    public List<Slot> Slots = new List<Slot>();
    public string CaptureMod = "Alt";
    public string CaptureKey = "1";
    public string PasteMod = "Alt";
    public string PasteKey = "2";
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
            if(d.ContainsKey("capture_mod")) CaptureMod=d["capture_mod"].ToString();
            if(d.ContainsKey("capture_key")) CaptureKey=d["capture_key"].ToString();
            if(d.ContainsKey("paste_mod")) PasteMod=d["paste_mod"].ToString();
            if(d.ContainsKey("paste_key")) PasteKey=d["paste_key"].ToString();
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
        doc["capture_mod"]=CaptureMod;
        doc["capture_key"]=CaptureKey;
        doc["paste_mod"]=PasteMod;
        doc["paste_key"]=PasteKey;
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
        doc["capture_mod"]="Alt";
        doc["capture_key"]="1";
        doc["paste_mod"]="Alt";
        doc["paste_key"]="2";
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
// SendInput structures (replaces deprecated keybd_event)
[StructLayout(LayoutKind.Sequential)]
struct KEYBDINPUT {
    public ushort wVk; public ushort wScan; public uint dwFlags;
    public uint time; public IntPtr dwExtraInfo;
}
[StructLayout(LayoutKind.Sequential)]
struct MOUSEINPUT {
    public int dx,dy; public uint mouseData,dwFlags,time; public IntPtr dwExtraInfo;
}
[StructLayout(LayoutKind.Explicit)]
struct INPUTUNION {
    [FieldOffset(0)] public MOUSEINPUT mi;
    [FieldOffset(0)] public KEYBDINPUT ki;
}
[StructLayout(LayoutKind.Sequential)]
struct INPUT {
    public uint type; public INPUTUNION u;
}

class App : Form {
    const string VER="v1.4.1";
    // ── Win32 APIs ──
    [DllImport("user32")] static extern bool RegisterHotKey(IntPtr h,int id,uint mod,uint vk);
    [DllImport("user32")] static extern bool UnregisterHotKey(IntPtr h,int id);
    [DllImport("user32")] static extern bool DestroyIcon(IntPtr handle);
    [DllImport("user32")] static extern uint SendInput(uint n, INPUT[] inputs, int size);
    // Event-driven clipboard (replaces Timer polling)
    [DllImport("user32")] static extern bool AddClipboardFormatListener(IntPtr hwnd);
    [DllImport("user32")] static extern bool RemoveClipboardFormatListener(IntPtr hwnd);
    const int WM_CLIPBOARDUPDATE = 0x031D;
    // Window-aware capture
    [DllImport("user32")] static extern IntPtr GetForegroundWindow();
    [DllImport("user32",CharSet=CharSet.Auto)] static extern int GetWindowText(IntPtr h, System.Text.StringBuilder sb, int max);
    // Capture window title keywords (configurable in JSON)
    List<string> captureWindowKeywords = new List<string>{"ORD","\u6aa2\u9a57","\u5831\u544a","HIS","EMR"};

    NotifyIcon tray; Config cfg = new Config();
    string labResult; string lastClip; DateTime ignoreUntil;
    const uint MOD_CTRL=2; const uint MOD_ALT=1; const int WM_HOTKEY=0x312;
    // Hotkey IDs: 0=settings, 1~4=output slots, 5=capture
    static readonly Dictionary<int,uint> HK = new Dictionary<int,uint>{
        {0,0x30},{1,0x31},{2,0x32},{3,0x33},{4,0x34}};
    const int HKID_CAPTURE=5;
    const int HKID_PASTE=6;

    // ── Multi-report merge buffer ──
    List<string> rawBuffer = new List<string>();
    System.Windows.Forms.Timer mergeTimer = new System.Windows.Forms.Timer{Interval=15000};

    Icon MakeIcon(Color c) {
        var bmp=new Bitmap(32,32); using(var g=Graphics.FromImage(bmp)){
            g.SmoothingMode=System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            g.Clear(Color.Transparent);
            // Soft status glow behind head
            using(var b=new SolidBrush(Color.FromArgb(35,c)))g.FillEllipse(b,0,3,31,28);
            // Big round face (chibi proportion)
            using(var b=new SolidBrush(Color.FromArgb(255,225,195)))g.FillEllipse(b,3,10,26,21);
            // Hair tufts (dark brown)
            using(var b=new SolidBrush(Color.FromArgb(90,55,30))){
                g.FillEllipse(b,3,11,7,8); g.FillEllipse(b,22,11,7,8);}
            // Doctor cap (white puffy dome)
            using(var b=new SolidBrush(Color.White)){
                g.FillEllipse(b,5,0,22,14);
                g.FillRectangle(b,4,7,24,6);}
            // Cross on cap (status color)
            using(var b=new SolidBrush(c)){
                g.FillRectangle(b,14,2,4,8);
                g.FillRectangle(b,11,4,10,4);}
            // Cap brim
            using(var p=new Pen(Color.FromArgb(180,200,200,200),0.8f))g.DrawLine(p,5,12,27,12);
            // Big kawaii eyes
            using(var b=new SolidBrush(Color.FromArgb(45,45,55))){
                g.FillEllipse(b,8,16,6,6); g.FillEllipse(b,18,16,6,6);}
            // Eye sparkle (big)
            using(var b=new SolidBrush(Color.White)){
                g.FillEllipse(b,9,16,3,3); g.FillEllipse(b,19,16,3,3);}
            // Eye sparkle (tiny)
            using(var b=new SolidBrush(Color.FromArgb(220,255,255,255))){
                g.FillEllipse(b,12,20,2,1); g.FillEllipse(b,22,20,2,1);}
            // Rosy blush cheeks
            using(var b=new SolidBrush(Color.FromArgb(65,255,100,110))){
                g.FillEllipse(b,4,22,8,4); g.FillEllipse(b,20,22,8,4);}
            // Happy cat smile :3
            using(var p=new Pen(Color.FromArgb(210,90,70),1.2f)){
                g.DrawArc(p,10,23,6,3,10,160);
                g.DrawArc(p,16,23,6,3,10,160);}
        }
        var hIcon=bmp.GetHicon(); var icon=Icon.FromHandle(hIcon);
        var clone=(Icon)icon.Clone(); DestroyIcon(hIcon); bmp.Dispose();
        return clone;
    }

    public App() {
        cfg.Load(); ShowInTaskbar=false; WindowState=FormWindowState.Minimized; Visible=false;
        tray=new NotifyIcon{Icon=MakeIcon(Color.Gold),Text="Lab Formatter "+VER,Visible=true};
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
        RegisterCaptureHotkey();
        RegisterPasteHotkey();

        lastClip=SafeGetClip(); ignoreUntil=DateTime.MinValue;
        // Wire up timer events (initialized at field level, handlers set once)
        mergeTimer.Tick+=(s,e)=>{mergeTimer.Stop(); FinalizeMerge();};
        previewTimer.Tick+=(s,e)=>{previewTimer.Stop();if(previewForm!=null&&!previewForm.IsDisposed)previewForm.Hide();};
        yellowTimer.Tick+=(s,e)=>{yellowTimer.Stop();
            SetTray(Color.Gold,labResult!=null?
                cfg.PasteMod+"+"+cfg.PasteKey+" \u8cbc\u4e0a | "+labResult.Substring(0,Math.Min(40,labResult.Length)):
                cfg.CaptureMod+"+"+cfg.CaptureKey+" \u64f7\u53d6");};
        // Event-driven clipboard monitoring (no more Timer polling)
        AddClipboardFormatListener(Handle);
    }

    // ── Safe clipboard with retry (prevents ExternalException crashes) ──
    string SafeGetClip(){
        for(int retry=0;retry<3;retry++){
            try{if(Clipboard.ContainsText()) return Clipboard.GetText();}
            catch{Thread.Sleep(30);}
        }
        return "";
    }
    bool SafeSetClip(string t){
        for(int retry=0;retry<3;retry++){
            try{Clipboard.SetText(t);return true;}
            catch{Thread.Sleep(30);}
        }
        return false;
    }

    // Called by WM_CLIPBOARDUPDATE (event-driven, replaces Timer polling)
    void OnClipboardChanged() {
        var t=SafeGetClip();
        if(t!=""&&t!=lastClip){lastClip=t;
            if(DateTime.Now<ignoreUntil)return;
            if(!Lab.IsLabData(t))return;
            rawBuffer.Add(t);
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
    System.Windows.Forms.Timer previewTimer = new System.Windows.Forms.Timer();
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
            previewTitle.Text="\u2713 "+count+"\u7b46\u5831\u544a\u5df2\u8f49\u63db\uff0c15\u79d2\u5167\u53ef\u7e7c\u7e8c\u64f7\u53d6 | "+cfg.PasteMod+"+"+cfg.PasteKey+" \u8cbc\u4e0a";
        else
            previewTitle.Text="\u2713 \u5831\u544a\u5df2\u8f49\u63db\uff0c\u6309 "+cfg.PasteMod+"+"+cfg.PasteKey+" \u8cbc\u4e0a";

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
        previewTimer.Interval=hideDelay;
        previewTimer.Stop(); previewTimer.Start();
    }
    // Hide preview when Ctrl+3 pastes
    void HidePreview(){
        if(previewForm!=null&&!previewForm.IsDisposed&&previewForm.Visible) previewForm.Hide();
        previewTimer.Stop();
    }
    void SetTray(Color c,string tip){
        var old=tray.Icon; tray.Icon=MakeIcon(c); if(old!=null)old.Dispose();
        tray.Text=tip.Length>63?tip.Substring(0,63):tip;}
    System.Windows.Forms.Timer yellowTimer = new System.Windows.Forms.Timer{Interval=3000};
    void DelayYellow(){
        yellowTimer.Stop(); yellowTimer.Start();
    }
    void WriteClip(string t){ignoreUntil=DateTime.Now.AddMilliseconds(800);
        SafeSetClip(t); lastClip=t;}

    // ── SendInput helpers (replaces deprecated keybd_event) ──
    static void SendKeys(params ushort[] vks){
        var inputs=new List<INPUT>();
        // Key down
        foreach(var vk in vks){
            var inp=new INPUT(); inp.type=1; inp.u.ki.wVk=vk; inp.u.ki.dwFlags=0;
            inputs.Add(inp);
        }
        // Key up (reverse order)
        for(int i=vks.Length-1;i>=0;i--){
            var inp=new INPUT(); inp.type=1; inp.u.ki.wVk=vks[i]; inp.u.ki.dwFlags=2;
            inputs.Add(inp);
        }
        SendInput((uint)inputs.Count, inputs.ToArray(), Marshal.SizeOf(typeof(INPUT)));
    }
    static void ReleaseAllModifiers(){
        // Release Ctrl(0x11), Alt(0x12), Shift(0x10) to prevent ghost modifiers
        var inputs=new INPUT[3];
        inputs[0]=new INPUT(); inputs[0].type=1; inputs[0].u.ki.wVk=0x12; inputs[0].u.ki.dwFlags=2;
        inputs[1]=new INPUT(); inputs[1].type=1; inputs[1].u.ki.wVk=0x11; inputs[1].u.ki.dwFlags=2;
        inputs[2]=new INPUT(); inputs[2].type=1; inputs[2].u.ki.wVk=0x10; inputs[2].u.ki.dwFlags=2;
        SendInput(3, inputs, Marshal.SizeOf(typeof(INPUT)));
    }

    void SimCtrlV(){
        ThreadPool.QueueUserWorkItem(delegate{
            Thread.Sleep(50);
            ReleaseAllModifiers(); Thread.Sleep(30);
            SendKeys(0x11, 0x56); // Ctrl+V
        });
    }
    void SimCtrlAC(){
        ThreadPool.QueueUserWorkItem(delegate{
            Thread.Sleep(50);
            ReleaseAllModifiers(); Thread.Sleep(30);
            SendKeys(0x11, 0x41); // Ctrl+A
            Thread.Sleep(80);
            SendKeys(0x11, 0x43); // Ctrl+C
        });
    }

    // ── Window-aware capture: only trigger in ORD/HIS windows ──
    bool IsCaptureAllowed(){
        var hwnd=GetForegroundWindow();
        if(hwnd==IntPtr.Zero) return false;
        var sb=new System.Text.StringBuilder(512);
        GetWindowText(hwnd,sb,512);
        var title=sb.ToString();
        if(title=="") return true; // Unknown window, allow (fallback)
        foreach(var kw in captureWindowKeywords)
            if(title.IndexOf(kw,StringComparison.OrdinalIgnoreCase)>=0) return true;
        return false;
    }
    void RegisterCaptureHotkey(){
        UnregisterHotKey(Handle,HKID_CAPTURE);
        uint mod=VK.ModFlag(cfg.CaptureMod);
        uint vk=VK.Code(cfg.CaptureKey);
        RegisterHotKey(Handle,HKID_CAPTURE,mod,vk);
    }
    void RegisterPasteHotkey(){
        UnregisterHotKey(Handle,HKID_PASTE);
        uint mod=VK.ModFlag(cfg.PasteMod);
        uint vk=VK.Code(cfg.PasteKey);
        RegisterHotKey(Handle,HKID_PASTE,mod,vk);
    }

    protected override void WndProc(ref Message m) {
        // Event-driven clipboard monitoring
        if(m.Msg==WM_CLIPBOARDUPDATE){ OnClipboardChanged(); base.WndProc(ref m); return; }
        if(m.Msg==WM_HOTKEY){int id=(int)m.WParam;
            if(id==0){ShowSettings();return;}
            if(id==HKID_CAPTURE){
                // Window-aware: only capture in ORD/HIS windows
                if(IsCaptureAllowed()) SimCtrlAC();
                return;
            }
            if(id==HKID_PASTE){
                // Same as lab-type paste
                if(rawBuffer.Count>0){
                    var combined=string.Join("\n",rawBuffer);
                    labResult=Lab.Convert(combined,cfg.GetLabItems());
                    rawBuffer.Clear();
                    if(mergeTimer!=null) mergeTimer.Stop();
                }
                if(labResult!=null){WriteClip(labResult);SimCtrlV();
                HidePreview();SetTray(Color.LimeGreen,"\u2713 \u5df2\u8cbc\u4e0a");DelayYellow();}
                else{SetTray(Color.Red,"\u8acb\u5148\u64f7\u53d6\u5831\u544a");DelayYellow();}
                return;
            }
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
                    var clip=SafeGetClip();WriteClip(s.Text.Replace("{clipboard}",clip));SimCtrlV();}
            }catch{}}
        base.WndProc(ref m);
    }
    protected override void OnFormClosed(FormClosedEventArgs e){
        RemoveClipboardFormatListener(Handle);
        foreach(var k in HK.Keys) UnregisterHotKey(Handle,k);
        UnregisterHotKey(Handle,HKID_CAPTURE);
        UnregisterHotKey(Handle,HKID_PASTE);
        mergeTimer.Stop(); mergeTimer.Dispose();
        yellowTimer.Stop(); yellowTimer.Dispose();
        previewTimer.Stop(); previewTimer.Dispose();
        if(previewForm!=null&&!previewForm.IsDisposed) previewForm.Close();
        tray.Visible=false; tray.Icon.Dispose(); tray.Dispose();
        base.OnFormClosed(e);}
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

        var lbl=new Label{Text=cfg.CaptureMod+"+"+cfg.CaptureKey+" \u64f7\u53d6 | "+cfg.PasteMod+"+"+cfg.PasteKey+" \u8cbc\u4e0a | Ctrl+0 \u8a2d\u5b9a",
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
        var helpBtn=new Button{Text="\u4f7f\u7528\u8aaa\u660e",Left=fw/2+10,Top=sy+2,Width=110,Height=30,
            BackColor=Color.FromArgb(33,150,243),ForeColor=Color.White,
            Font=new Font("Microsoft JhengHei UI",10)};
        helpBtn.Click+=(s,e)=>OpenReadme();
        f.Controls.Add(saveBtn);f.Controls.Add(helpBtn);

        // ── Item checkboxes ──
        int iy=sy+40;
        var itemGb=new GroupBox{Text="\u8f38\u51fa\u9805\u76ee\u52fe\u9078\uff08\u53d6\u6d88\u52fe\u9078 = \u4e0d\u7d0d\u5165\u6574\u7406\uff09",
            Left=10,Top=iy,Width=fw-20,Height=150,Font=new Font("Microsoft JhengHei UI",9)};
        f.Controls.Add(itemGb);
        var chkItems=new Dictionary<string,CheckBox>();
        var labEnabled=cfg.GetLabItems();
        // Toggle all button
        bool toggleState=true;
        var toggleBtn=new Button{Text="\u5168\u4e0d\u9078",Left=itemGb.Width-80,Top=16,Width=68,Height=22,
            Font=new Font("Microsoft JhengHei UI",8)};
        toggleBtn.Click+=(s,e)=>{
            toggleState=!toggleState;
            foreach(var kv in chkItems) kv.Value.Checked=toggleState;
            toggleBtn.Text=toggleState?"\u5168\u4e0d\u9078":"\u5168\u9078";
        };
        itemGb.Controls.Add(toggleBtn);
        int cx=12,cy=18;
        foreach(var it in Lab.Items){
            var cb=new CheckBox{Text=it.Disp+"("+it.Key+")",Left=cx,Top=cy,Width=110,Height=20,
                Checked=labEnabled.Contains(it.Key),Font=new Font("Microsoft JhengHei UI",8)};
            chkItems[it.Key]=cb; itemGb.Controls.Add(cb);
            cx+=115; if(cx>fw-140){cx=12;cy+=22;}
        }

        // ── Capture hotkey row ──
        int chy=iy+itemGb.Height+6;
        f.Controls.Add(new Label{Text="\u64f7\u53d6\u9375\uff08\u5168\u9078+\u8907\u88fd\uff09:",
            Left=10,Top=chy+3,Width=130,AutoSize=false,
            Font=new Font("Microsoft JhengHei UI",9)});
        var capModCombo=new ComboBox{Left=142,Top=chy,Width=65,DropDownStyle=ComboBoxStyle.DropDownList,
            Font=new Font("Microsoft JhengHei UI",9)};
        capModCombo.Items.AddRange(new object[]{"Ctrl","Alt"});
        capModCombo.SelectedItem=cfg.CaptureMod;
        f.Controls.Add(capModCombo);
        f.Controls.Add(new Label{Text="+",Left=210,Top=chy+3,Width=14,AutoSize=false,
            Font=new Font("Microsoft JhengHei UI",9)});
        var capKeyBox=new TextBox{Left=226,Top=chy,Width=50,Text=cfg.CaptureKey,MaxLength=5,
            Font=new Font("Consolas",10),TextAlign=HorizontalAlignment.Center};
        f.Controls.Add(capKeyBox);
        f.Controls.Add(new Label{Text="\u76ee\u524d: "+cfg.CaptureMod+"+"+cfg.CaptureKey+" \u2192 \u81ea\u52d5\u5168\u9078+\u8907\u88fd",
            Left=286,Top=chy+3,Width=280,AutoSize=false,
            ForeColor=Color.Gray,Font=new Font("Microsoft JhengHei UI",8)});

        // ── Paste hotkey row ──
        int phy=chy+26;
        f.Controls.Add(new Label{Text="\u8cbc\u4e0a\u9375\uff08\u8cbc\u4e0a\u7d50\u679c\uff09:",
            Left=10,Top=phy+3,Width=130,AutoSize=false,
            Font=new Font("Microsoft JhengHei UI",9)});
        var pasteModCombo=new ComboBox{Left=142,Top=phy,Width=65,DropDownStyle=ComboBoxStyle.DropDownList,
            Font=new Font("Microsoft JhengHei UI",9)};
        pasteModCombo.Items.AddRange(new object[]{"Ctrl","Alt"});
        pasteModCombo.SelectedItem=cfg.PasteMod;
        f.Controls.Add(pasteModCombo);
        f.Controls.Add(new Label{Text="+",Left=210,Top=phy+3,Width=14,AutoSize=false,
            Font=new Font("Microsoft JhengHei UI",9)});
        var pasteKeyBox=new TextBox{Left=226,Top=phy,Width=50,Text=cfg.PasteKey,MaxLength=5,
            Font=new Font("Consolas",10),TextAlign=HorizontalAlignment.Center};
        f.Controls.Add(pasteKeyBox);
        f.Controls.Add(new Label{Text="\u76ee\u524d: "+cfg.PasteMod+"+"+cfg.PasteKey+" \u2192 \u8cbc\u4e0a\u6574\u7406\u7d50\u679c",
            Left=286,Top=phy+3,Width=280,AutoSize=false,
            ForeColor=Color.Gray,Font=new Font("Microsoft JhengHei UI",8)});

        saveBtn.Click+=(s,e)=>{
            for(int i=0;i<4;i++) slots[i].SaveTo(cfg.Slots[i]);
            // Save checked lab items
            var checkedKeys=new List<string>();
            foreach(var kv in chkItems) if(kv.Value.Checked) checkedKeys.Add(kv.Key);
            foreach(var sl in cfg.Slots) if(sl.Type=="lab") sl.LabItems=checkedKeys;
            // Save capture + paste hotkeys
            cfg.CaptureMod=capModCombo.Text;
            cfg.CaptureKey=capKeyBox.Text.Trim();
            if(cfg.CaptureKey=="") cfg.CaptureKey="1";
            cfg.PasteMod=pasteModCombo.Text;
            cfg.PasteKey=pasteKeyBox.Text.Trim();
            if(cfg.PasteKey=="") cfg.PasteKey="2";
            cfg.Save();
            RegisterCaptureHotkey();
            RegisterPasteHotkey();
            SetTray(Color.LimeGreen,"\u2713 \u5df2\u5132\u5b58");DelayYellow();
        };

        // JSON row
        int jy=phy+28;
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
        Splash.Show();
        Application.Run(new App());
    }
}

// ══════════ Splash Screen ══════════
static class Splash {
    public static void Show() {
        var f=new Form{
            FormBorderStyle=FormBorderStyle.None,
            StartPosition=FormStartPosition.CenterScreen,
            Size=new Size(420,200),
            BackColor=Color.FromArgb(16,24,40),
            ShowInTaskbar=false,
            TopMost=true,
            Opacity=0
        };
        // Round corners
        var path=new System.Drawing.Drawing2D.GraphicsPath();
        path.AddArc(0,0,20,20,180,90); path.AddArc(f.Width-20,0,20,20,270,90);
        path.AddArc(f.Width-20,f.Height-20,20,20,0,90); path.AddArc(0,f.Height-20,20,20,90,90);
        path.CloseFigure();
        f.Region=new Region(path);

        // Title
        var lbl1=new Label{Text="\u75c5\u6b77\u5c0f\u5e6b\u624b",
            ForeColor=Color.FromArgb(74,222,128),
            Font=new Font("Microsoft JhengHei UI",28,FontStyle.Bold),
            AutoSize=false,Width=420,Height=70,Top=30,
            TextAlign=ContentAlignment.MiddleCenter};
        // Subtitle
        var lbl2=new Label{Text="\u4e0a\u7dda\u4e2d...",
            ForeColor=Color.FromArgb(148,163,184),
            Font=new Font("Microsoft JhengHei UI",14),
            AutoSize=false,Width=420,Height=40,Top=100,
            TextAlign=ContentAlignment.MiddleCenter};
        // Version
        var lbl3=new Label{Text="Lab Data Formatter",
            ForeColor=Color.FromArgb(100,116,139),
            Font=new Font("Consolas",10),
            AutoSize=false,Width=420,Height=25,Top=150,
            TextAlign=ContentAlignment.MiddleCenter};
        f.Controls.AddRange(new Control[]{lbl1,lbl2,lbl3});

        int phase=0; int tick=0;
        int origW=f.Width, origH=f.Height;
        var cx=f.Left+origW/2; var cy=f.Top+origH/2;

        var timer=new System.Windows.Forms.Timer{Interval=30};
        timer.Tick+=delegate{
            tick++;
            if(phase==0){
                // Fade in (0~15 ticks = 0.45s)
                double p=(double)tick/15; if(p>1)p=1;
                f.Opacity=p;
                if(tick>=15){phase=1;tick=0;lbl2.Text="\u4e0a\u7dda\u5b8c\u6210 \u2713";
                    lbl2.ForeColor=Color.FromArgb(74,222,128);}
            } else if(phase==1){
                // Hold (0~50 ticks = 1.5s)
                if(tick>=50){phase=2;tick=0;
                    cx=f.Left+f.Width/2; cy=f.Top+f.Height/2;}
            } else if(phase==2){
                // Shrink + fade out (0~20 ticks = 0.6s)
                double p=(double)tick/20; if(p>1)p=1;
                f.Opacity=1.0-p;
                int w=(int)(origW*(1.0-p*0.6));
                int h=(int)(origH*(1.0-p*0.6));
                if(w<10)w=10; if(h<10)h=10;
                f.SetBounds(cx-w/2,cy-h/2,w,h);
                // Update region for new size
                var rp=new System.Drawing.Drawing2D.GraphicsPath();
                int r2=Math.Max((int)(10*(1.0-p)),2);
                rp.AddArc(0,0,r2*2,r2*2,180,90); rp.AddArc(w-r2*2,0,r2*2,r2*2,270,90);
                rp.AddArc(w-r2*2,h-r2*2,r2*2,r2*2,0,90); rp.AddArc(0,h-r2*2,r2*2,r2*2,90,90);
                rp.CloseFigure();
                f.Region=new Region(rp);
                if(tick>=20){timer.Stop();f.Close();}
            }
        };
        f.Shown+=delegate{timer.Start();};
        f.ShowDialog();
    }
}
