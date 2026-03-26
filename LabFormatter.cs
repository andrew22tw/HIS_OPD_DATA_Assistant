// Lab Data Formatter v2.0
// Author: 吳岳霖醫師 (DAL93@tpech.gov.tw)
// Compile: csc.exe /target:winexe LabFormatter.cs
// v2.0: NHI Cloud lab + medication support, Apple-style UI
// Architecture: Event-driven clipboard (WM_CLIPBOARDUPDATE), SendInput

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
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
        i++; var sb=new StringBuilder();
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

// ══════════ Apple-style Theme ══════════
static class Theme {
    // Dark mode colors (iOS-inspired)
    public static readonly Color Bg = Color.FromArgb(28, 28, 30);           // #1C1C1E
    public static readonly Color Surface = Color.FromArgb(44, 44, 46);      // #2C2C2E
    public static readonly Color Surface2 = Color.FromArgb(58, 58, 60);     // #3A3A3C
    public static readonly Color Border = Color.FromArgb(68, 68, 70);       // #444446
    public static readonly Color Accent = Color.FromArgb(52, 199, 89);      // #34C759 (iOS green)
    public static readonly Color AccentDim = Color.FromArgb(30, 52, 199, 89);
    public static readonly Color Text = Color.FromArgb(245, 245, 247);      // #F5F5F7
    public static readonly Color Text2 = Color.FromArgb(142, 142, 147);     // #8E8E93
    public static readonly Color Text3 = Color.FromArgb(99, 99, 102);       // #636366
    public static readonly Color Red = Color.FromArgb(255, 69, 58);         // iOS red
    public static readonly Color Blue = Color.FromArgb(10, 132, 255);       // iOS blue
    public static readonly Color Gold = Color.FromArgb(255, 214, 10);       // iOS yellow
    public static readonly Color CardBg = Color.FromArgb(38, 38, 40);

    public static Font Title = new Font("Microsoft JhengHei UI", 11, FontStyle.Bold);
    public static Font Body = new Font("Microsoft JhengHei UI", 9);
    public static Font Mono = new Font("Consolas", 10);
    public static Font MonoSmall = new Font("Consolas", 9);
    public static Font Small = new Font("Microsoft JhengHei UI", 8);

    public static GraphicsPath RoundRect(Rectangle r, int radius) {
        var p = new GraphicsPath();
        int d = radius * 2;
        p.AddArc(r.X, r.Y, d, d, 180, 90);
        p.AddArc(r.Right - d, r.Y, d, d, 270, 90);
        p.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
        p.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
        p.CloseFigure();
        return p;
    }

    public static void SetRoundRegion(Form f, int radius) {
        f.Region = new Region(RoundRect(new Rectangle(0, 0, f.Width, f.Height), radius));
    }
}

// ══════════ Pill Button ══════════
class PillButton : Button {
    public Color AccentColor = Theme.Accent;
    public bool IsPrimary = true;
    public PillButton() {
        FlatStyle = FlatStyle.Flat;
        FlatAppearance.BorderSize = 0;
        FlatAppearance.MouseOverBackColor = Color.Transparent;
        FlatAppearance.MouseDownBackColor = Color.Transparent;
        Cursor = Cursors.Hand;
        Font = Theme.Body;
        ForeColor = Color.White;
        BackColor = Color.Transparent;
    }
    protected override void OnPaint(PaintEventArgs e) {
        var g = e.Graphics;
        g.SmoothingMode = SmoothingMode.AntiAlias;
        var r = new Rectangle(0, 0, Width - 1, Height - 1);
        using (var path = Theme.RoundRect(r, Height / 2)) {
            if (IsPrimary)
                using (var b = new SolidBrush(AccentColor)) g.FillPath(b, path);
            else
                using (var p = new Pen(Theme.Border, 1.5f)) g.DrawPath(p, path);
        }
        TextRenderer.DrawText(g, Text, Font, r, ForeColor,
            TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
    }
}

// ══════════ Card Panel ══════════
class CardPanel : Panel {
    public string Title = "";
    public CardPanel() {
        BackColor = Color.Transparent;
        DoubleBuffered = true;
        Padding = new Padding(16, 36, 16, 12);
    }
    protected override void OnPaint(PaintEventArgs e) {
        var g = e.Graphics;
        g.SmoothingMode = SmoothingMode.AntiAlias;
        var r = new Rectangle(0, 0, Width - 1, Height - 1);
        using (var path = Theme.RoundRect(r, 12))
        using (var b = new SolidBrush(Theme.CardBg)) g.FillPath(b, path);
        if (Title != "") {
            using (var b = new SolidBrush(Theme.Text))
                g.DrawString(Title, Theme.Title, b, 16, 12);
        }
    }
}

// ══════════ Lab Parsing (ORD system) ══════════
static class Lab {
    public class Item {
        public string Key, Pattern, Disp; public bool On;
        public Regex Rx;
    }
    static readonly Regex RxFlag = new Regex(@"(?:^|\s)[HL]{1,2}\s*([\d]+\.?\d*)", RegexOptions.Compiled);
    static readonly Regex RxVal = new Regex(@"\s{2,}([\d]+\.?\d*)", RegexOptions.Compiled);

    public static List<Item> Items = new List<Item> {
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
    static readonly Regex RxDate = new Regex("採檢時間[：:]?\\s*(\\d{2,3})/(\\d{2})/(\\d{2})", RegexOptions.Compiled);
    public static string ExtractDate(string text) {
        var m = RxDate.Match(text);
        return m.Success ? m.Groups[1].Value + m.Groups[2].Value : "";
    }
    public static bool IsLabData(string text) {
        var kws = new[]{"BUN","mg/dl","mg/dL","mEq/L","U/L","mg/L","Glu","Cholesterol",
            "LDL","HDL","ALT","AST","Hemolysis","Lipemia","Icterus","Cr ",
            "採檢時間","檢驗科","結果值","參考值","檢體","報告者"};
        return kws.Count(k => text.Contains(k)) >= 3;
    }

    // Format values dict into condensed string (reusable by CloudLab)
    public static string Format(string date, Dictionary<string,string> vals, HashSet<string> enabled, string dateSuffix=" ") {
        var roundKeys = new[]{"BUN","HCO3","Iron","TIBC","Ferritin"};
        foreach(var rk in roundKeys) {
            if(vals.ContainsKey(rk)){double d;if(double.TryParse(vals[rk],out d))vals[rk]=((int)Math.Round(d)).ToString();}
        }
        var parts = new List<string>();
        var used = new HashSet<string>();
        foreach (var it in Items) {
            if (used.Contains(it.Key)||!enabled.Contains(it.Key)||!vals.ContainsKey(it.Key)) continue;
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
        int indent = 0;
        if (date!="") { var prefix=date+dateSuffix; r=prefix+r; indent=prefix.Length; }
        // Word-wrap at 80 chars, indent continuation to align under first data char
        if (r.Length>80) {
            var pad = new string(' ', indent);
            var sb = new StringBuilder();
            int col=0;
            var tokens=r.Split(',');
            for(int ti=0;ti<tokens.Length;ti++){
                var tok=tokens[ti]+(ti<tokens.Length-1?",":"");
                if(col==0){sb.Append(tok);col=tok.Length;}
                else if(col+tok.Length>80){
                    sb.Append("\n").Append(pad);col=indent;sb.Append(tok);col+=tok.Length;}
                else{sb.Append(tok);col+=tok.Length;}
            }
            r=sb.ToString();
        }
        return r;
    }

    public static string Convert(string raw, HashSet<string> enabled) {
        var date = ExtractDate(raw);
        var vals = new Dictionary<string,string>();
        var urineKeys = new HashSet<string>{"UAC","UPC"};
        bool inUrine = false;
        foreach (var line in raw.Split('\n')) {
            if (line.Contains("檢體")) { inUrine = line.Contains("尿液"); }
            foreach (var it in Items) {
                if (vals.ContainsKey(it.Key)) continue;
                bool isUrineItem = urineKeys.Contains(it.Key);
                if (isUrineItem && !inUrine) continue;
                if (!isUrineItem && inUrine) continue;
                var v = ExtractVal(line, it.Rx);
                if (v != null) vals[it.Key] = v;
            }
        }
        return Format(date, vals, enabled);
    }
}

// ══════════ Cloud Lab Parsing (NHI MedCloud) ══════════
static class CloudLab {
    // Map cloud test names → Lab.Items keys
    // This covers common names uploaded by different hospitals
    static readonly Dictionary<string, string> NameMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) {
        // Glucose
        {"Glucose AC", "GluAC"}, {"Glu AC", "GluAC"}, {"Glu.AC", "GluAC"},
        {"飯前血糖", "GluAC"}, {"空腹血糖", "GluAC"}, {"Fasting glucose", "GluAC"},
        {"Glucose (AC)", "GluAC"}, {"Sugar(AC)", "GluAC"},
        // HbA1c
        {"HbA1c", "HbA1c"}, {"HbA1C", "HbA1c"}, {"Glycated Hb", "HbA1c"},
        {"糖化血色素", "HbA1c"},
        // BUN
        {"BUN", "BUN"}, {"Blood urea nitrogen", "BUN"}, {"Urea nitrogen", "BUN"},
        {"尿素氮", "BUN"},
        // Creatinine
        {"Creatinine", "Cr"}, {"Cr", "Cr"}, {"Creatinine(B)", "Cr"},
        {"肌酸酐", "Cr"},
        // eGFR
        {"eGFR", "eGFR"}, {"estimated GFR", "eGFR"}, {"estimated Ccr(MDRD)", "eGFR"},
        {"eGFR(MDRD)", "eGFR"}, {"eCCr", "eGFR"},
        {"腎絲球過濾率", "eGFR"},
        // Na
        {"Na", "Na"}, {"Sodium", "Na"}, {"鈉", "Na"},
        // K
        {"K", "K"}, {"Potassium", "K"}, {"鉀", "K"},
        // HCO3
        {"HCO3", "HCO3"}, {"Bicarbonate", "HCO3"}, {"TCO2", "HCO3"},
        // CRP
        {"CRP", "CRP"}, {"C-Reactive Protein", "CRP"},
        // ALT
        {"ALT", "ALT"}, {"GPT", "ALT"}, {"ALT(GPT)", "ALT"}, {"SGPT", "ALT"},
        {"麩丙酮酸轉氨基酶", "ALT"},
        // AST
        {"AST", "AST"}, {"GOT", "AST"}, {"AST(GOT)", "AST"}, {"SGOT", "AST"},
        {"麩胺酸苯醋酸轉氨基酶", "AST"},
        // Albumin
        {"Albumin", "Alb"}, {"Alb", "Alb"}, {"白蛋白", "Alb"},
        // Cholesterol
        {"Cholesterol", "Chol"}, {"T-Cholesterol", "Chol"}, {"Total Cholesterol", "Chol"},
        {"總膽固醇", "Chol"}, {"Chol", "Chol"},
        // TG
        {"TG", "TG"}, {"Triglyceride", "TG"}, {"三酸甘油酯", "TG"},
        // LDL
        {"LDL-C", "LDL"}, {"LDL", "LDL"}, {"LDL Cholesterol", "LDL"},
        {"低密度脂蛋白", "LDL"},
        // HDL
        {"HDL-C", "HDL"}, {"HDL", "HDL"}, {"HDL Cholesterol", "HDL"},
        {"高密度脂蛋白", "HDL"},
        // UA
        {"Uric acid", "UA"}, {"UA", "UA"}, {"Uric Acid", "UA"}, {"尿酸", "UA"},
        // Ca
        {"Ca", "Ca"}, {"Calcium", "Ca"}, {"鈣", "Ca"},
        // P
        {"P", "P"}, {"Phosphorus", "P"}, {"Phosphate", "P"}, {"IP", "P"}, {"磷", "P"},
        // Iron
        {"Iron", "Iron"}, {"Fe", "Iron"}, {"鐵", "Iron"},
        // TIBC
        {"TIBC", "TIBC"},
        // Ferritin
        {"Ferritin", "Ferritin"}, {"鐵蛋白", "Ferritin"},
        // WBC
        {"WBC", "WBC"}, {"White Blood Cell", "WBC"}, {"白血球", "WBC"},
        // PLT
        {"Platelet", "PLT"}, {"PLT", "PLT"}, {"血小板", "PLT"},
        // T-Bil
        {"T-Bil", "TBil"}, {"T-Bilirubin", "TBil"}, {"Total Bilirubin", "TBil"},
        {"總膽紅素", "TBil"},
        // Mg
        {"Mg", "Mg"}, {"Magnesium", "Mg"}, {"鎂", "Mg"},
        // Hb
        {"Hb", "Hb"}, {"Hgb", "Hb"}, {"Hemoglobin", "Hb"}, {"血色素", "Hb"},
        // ACR
        {"UACR", "UAC"}, {"ACR", "UAC"}, {"Urine ACR", "UAC"},
        {"Albumin/Creatinine Ratio", "UAC"}, {"A/C Ratio", "UAC"},
        // PCR
        {"UPCR", "UPC"}, {"PCR", "UPC"}, {"Urine PCR", "UPC"},
        {"Protein/Creatinine Ratio", "UPC"}, {"P/C Ratio", "UPC"},
    };

    // Detect NHI cloud lab data (tab-separated with specific headers)
    public static bool IsCloudLabData(string text) {
        if (!text.Contains("\t")) return false;
        var headers = new[]{"檢驗日期","檢驗項目","檢驗結果"};
        int found = headers.Count(h => text.Contains(h));
        return found >= 2;
    }

    // Parse tab-separated table and convert to condensed format
    public static string Convert(string text, HashSet<string> enabled) {
        var lines = text.Split(new[]{'\n','\r'}, StringSplitOptions.RemoveEmptyEntries);
        if (lines.Length < 2) return null;

        // Find header row and build column map
        int headerIdx = -1;
        var colMap = new Dictionary<string, int>();
        for (int i = 0; i < Math.Min(5, lines.Length); i++) {
            if (lines[i].Contains("檢驗日期") || lines[i].Contains("檢驗項目")) {
                headerIdx = i;
                var cols = lines[i].Split('\t');
                for (int c = 0; c < cols.Length; c++) {
                    var h = cols[c].Trim();
                    if (h != "") colMap[h] = c;
                }
                break;
            }
        }
        if (headerIdx < 0) return null;

        // Required columns
        int dateCol = colMap.ContainsKey("檢驗日期") ? colMap["檢驗日期"] : -1;
        int nameCol = colMap.ContainsKey("檢驗項目") ? colMap["檢驗項目"] :
                      colMap.ContainsKey("醫令名稱") ? colMap["醫令名稱"] : -1;
        int resultCol = colMap.ContainsKey("檢驗結果") ? colMap["檢驗結果"] : -1;
        if (dateCol < 0 || resultCol < 0) return null;

        // Group data by date
        var byDate = new Dictionary<string, Dictionary<string, string>>();
        for (int i = headerIdx + 1; i < lines.Length; i++) {
            var cells = lines[i].Split('\t');
            if (cells.Length <= Math.Max(dateCol, Math.Max(nameCol < 0 ? 0 : nameCol, resultCol))) continue;

            var date = cells[dateCol].Trim();
            var testName = nameCol >= 0 && nameCol < cells.Length ? cells[nameCol].Trim() : "";
            var result = cells[resultCol].Trim();

            if (date == "" || result == "") continue;

            // Try to parse result as numeric value
            var valMatch = Regex.Match(result, @"([\d]+\.?\d*)");
            if (!valMatch.Success) continue;
            var val = valMatch.Groups[1].Value;

            // Map test name to Lab key
            string labKey = MapTestName(testName);
            if (labKey == null) continue;

            if (!byDate.ContainsKey(date))
                byDate[date] = new Dictionary<string, string>();
            if (!byDate[date].ContainsKey(labKey))
                byDate[date][labKey] = val;
        }

        if (byDate.Count == 0) return null;

        // Convert each date group to condensed format
        var results = new List<string>();
        foreach (var kv in byDate.OrderByDescending(x => x.Key)) {
            var dateStr = FormatCloudDate(kv.Key);
            var r = Lab.Format(dateStr, kv.Value, enabled, "(其他)");
            if (r != null) results.Add(r);
        }

        return results.Count > 0 ? string.Join("\n", results) : null;
    }

    static string MapTestName(string name) {
        if (string.IsNullOrEmpty(name)) return null;
        // Exact match first
        if (NameMap.ContainsKey(name)) return NameMap[name];
        // Fuzzy: trim, ignore case, try contains
        var trimmed = name.Trim();
        if (NameMap.ContainsKey(trimmed)) return NameMap[trimmed];
        // Try partial match
        foreach (var kv in NameMap) {
            if (trimmed.IndexOf(kv.Key, StringComparison.OrdinalIgnoreCase) >= 0)
                return kv.Value;
        }
        return null;
    }

    // Convert cloud date (e.g., "2024/03/15" or "113/03/15") to YYYMM format
    static string FormatCloudDate(string d) {
        var parts = d.Split('/');
        if (parts.Length >= 2) {
            int yr;
            if (int.TryParse(parts[0], out yr)) {
                // If western year (>200), convert to ROC
                if (yr > 200) yr -= 1911;
                return yr.ToString() + parts[1];
            }
        }
        return "";
    }
}

// ══════════ Cloud Medication Parsing (NHI MedCloud) ══════════
// Handles BOTH tab-separated AND continuous text (no delimiters) from clipboard
static class CloudMed {
    static readonly Dictionary<string, int> FreqMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase) {
        {"QD",1},{"QDP",1},{"QAM",1},{"QPM",1},{"QN",1},{"HS",1},{"HSP",1},{"DAILY",1},{"STAT",1},{"ST",1},
        {"BID",2},{"BIDP",2},{"Q12H",2},
        {"TID",3},{"TIDP",3},{"Q8H",3},
        {"QID",4},{"QIDP",4},{"Q6H",4},
        {"Q4H",6},{"Q2H",12},
    };
    static readonly Regex RxTablet = new Regex(@"\b(tablets?|f\.?c\.?\s*tablets?|film[\s-]?coated\s*(?:scored\s*)?tablets?|prolonged\s*release\s*tablets?)\b", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    static readonly Regex RxCapsule = new Regex(@"\b(ENTERIC-MICROENCAPSULATED|capsules?|cap\.?)\b", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    static readonly Regex RxDoseMg = new Regex(@"(\d+(?:\.\d+)?(?:\s*/\s*\d+(?:\.\d+)?){0,2})\s*mg\b(?!\s*/)", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    static readonly Regex RxComplexDose = new Regex(@"\((\d+(?:\.\d+)?(?:/\d+(?:\.\d+)?){0,2})\)\s*(?:MG|MCG|G|ML)(?!\s*/)", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    // ROC date pattern: 3-digit year / 2-digit month / 2-digit day
    static readonly Regex RxRocDate = new Regex(@"(\d{3}/\d{2}/\d{2})", RegexOptions.Compiled);
    // Row splitter: item number followed by ROC date
    static readonly Regex RxRowStart = new Regex(@"(?:^|\D)(\d{1,3})(\d{3}/\d{2}/\d{2})", RegexOptions.Compiled);
    // Frequency patterns at end of drug line (before route+numbers)
    static readonly string FreqPattern = @"(?:STAT|QDPC|BIDPC|TIDPC|QIDPC|QDP|QAM|QPM|QD|BID|TID|QID|Q\d+H|HSPC|HSP|HS|QN|QOD|TIW|BIW|QW|DAILY|PRN)";
    // Route patterns
    static readonly string RoutePattern = @"(?:PO|EXT|IV|IM|SC|IH|TOP|SL|RECT|INH|SKIN|OPH|NASAL|OTIC)";
    // Tail pattern: frequency + optional route + numbers + optional Y
    static readonly Regex RxTail = new Regex(
        @"(" + FreqPattern + @")(" + RoutePattern + @")?(\d+?)(\d{1,3}?)(\d{1,2})(Y?)$",
        RegexOptions.Compiled);
    // Simpler tail for short number strings (e.g., STAT220)
    static readonly Regex RxTailSimple = new Regex(
        @"(" + FreqPattern + @")(" + RoutePattern + @")?(\d+)(Y?)$",
        RegexOptions.Compiled);
    // ICD-10 code in diagnosis
    static readonly Regex RxICD = new Regex(@"([A-Z]\d{2,3}(?:\.\d+)?)\s*$", RegexOptions.Compiled);

    public class MedGroup {
        public string Date, Source, Diagnosis, DiagCode;
        public List<MedItem> Medicines = new List<MedItem>();
    }
    public class MedItem {
        public string Name, Ingredient, Usage, Days, TotalQty;
    }

    public static bool IsCloudMedData(string text) {
        // Tab-separated format
        if (text.Contains("\t")) {
            var headers = new[]{"就醫日期","藥品名稱","用法用量","給藥日數"};
            if (headers.Count(h => text.Contains(h)) >= 3) return true;
        }
        // Continuous text format: has header keywords + ROC dates + drug-like patterns
        if (text.Contains("藥品名稱") && text.Contains("就醫日期") && text.Contains("用法用量")) {
            if (RxRocDate.IsMatch(text)) return true;
        }
        return false;
    }

    public static string Convert(string text) {
        // Try tab-separated first
        if (text.Contains("\t") && text.Contains("藥品名稱")) {
            var r = ConvertTabSeparated(text);
            if (r != null) return r;
        }
        // Fall back to continuous text parsing
        return ConvertContinuous(text);
    }

    // ── Tab-separated parser ──
    static int ColIdx(Dictionary<string,int> m, string name) { return m.ContainsKey(name) ? m[name] : -1; }
    static string CellAt(string[] cells, int col) { return col >= 0 && col < cells.Length ? cells[col].Trim() : ""; }

    static string ConvertTabSeparated(string text) {
        var lines = text.Split(new[]{'\n','\r'}, StringSplitOptions.RemoveEmptyEntries);
        if (lines.Length < 2) return null;
        int headerIdx = -1;
        var colMap = new Dictionary<string, int>();
        for (int i = 0; i < Math.Min(5, lines.Length); i++) {
            if (lines[i].Contains("藥品名稱") && lines[i].Contains("就醫日期")) {
                headerIdx = i;
                var cols = lines[i].Split('\t');
                for (int c = 0; c < cols.Length; c++) {
                    var h = cols[c].Trim();
                    if (h != "") colMap[h] = c;
                }
                break;
            }
        }
        if (headerIdx < 0 || colMap.Count < 4) return null;

        int cDate=ColIdx(colMap,"就醫日期"), cSrc=ColIdx(colMap,"來源"), cDiag=ColIdx(colMap,"主診斷"),
            cName=ColIdx(colMap,"藥品名稱"), cUsage=ColIdx(colMap,"用法用量"),
            cDays=ColIdx(colMap,"給藥日數");
        if (cName < 0 || cDate < 0) return null;

        var groups = new Dictionary<string, MedGroup>();
        var seen = new HashSet<string>();
        for (int i = headerIdx + 1; i < lines.Length; i++) {
            var cells = lines[i].Split('\t');
            var date = CellAt(cells, cDate);
            var medName = CellAt(cells, cName);
            if (date == "" || medName == "") continue;
            var source = CellAt(cells, cSrc);
            var diag = CellAt(cells, cDiag);
            var usage = CellAt(cells, cUsage);
            var days = CellAt(cells, cDays);
            source = Regex.Replace(source, @"[\d\n\r]", "").Trim();
            string diagText = diag, diagCode = "";
            var dm = RxICD.Match(diag);
            if (dm.Success) { diagCode = dm.Groups[1].Value; diagText = diag.Substring(0, dm.Index).Trim(); }
            var gk = date+"_"+source+"_"+diagCode;
            var mk = date+"_"+medName;
            if (!groups.ContainsKey(gk))
                groups[gk] = new MedGroup{Date=date,Source=source,Diagnosis=diagText,DiagCode=diagCode};
            if (!seen.Contains(mk)) {
                seen.Add(mk);
                groups[gk].Medicines.Add(new MedItem{Name=medName,Usage=usage,Days=days});
            }
        }
        return FormatGroups(groups.Values.OrderByDescending(g=>g.Date).ToList());
    }

    // ── Continuous text parser (no tab delimiters) ──
    static string ConvertContinuous(string text) {
        // Split into individual medicine rows using ROC date as anchor
        // Pattern: (itemNumber)(ROC date)(rest of row until next item or end)
        var allMatches = RxRowStart.Matches(text);
        if (allMatches.Count == 0) return null;

        var groups = new Dictionary<string, MedGroup>();
        var seen = new HashSet<string>();

        for (int mi = 0; mi < allMatches.Count; mi++) {
            var m = allMatches[mi];
            var date = m.Groups[2].Value;  // e.g., "115/03/06"

            // Extract the rest of the row (until next row start or end of text)
            int rowStart = m.Index + m.Length;
            int rowEnd = (mi + 1 < allMatches.Count) ? allMatches[mi + 1].Index : text.Length;
            var rowText = text.Substring(rowStart, rowEnd - rowStart).Trim();

            // Skip pagination text like "Showing 1 to 50"
            if (rowText.Contains("Showing") || rowText.Contains("Previous") || rowText.Contains("Next")) continue;
            if (rowText.Length < 10) continue;

            // Parse source: starts with Chinese chars (hospital/pharmacy name)
            // Source typically ends before a diagnosis pattern (Chinese + ICD code)
            string source = "", diagText = "", diagCode = "", drugName = "", usage = "", days = "", ingredientName = "", totalQty = "";

            // Find ICD code in the row - it separates source+diagnosis from ATC+drug
            var icdMatch = Regex.Match(rowText, @"([A-Z]\d{2,3}(?:\.\d{1,2})?)");
            if (icdMatch.Success) {
                // Everything before ICD code area = source + diagnosis
                // Find the diagnosis: Chinese text before ICD code
                var beforeIcd = rowText.Substring(0, icdMatch.Index + icdMatch.Length);

                // Source is at the beginning (hospital name, may include 門診/急診/藥局 + numbers)
                // Diagnosis is Chinese text right before ICD code
                // Include Chinese commas, spaces, punctuation in diagnosis text
                var diagIcdMatch = Regex.Match(beforeIcd, @"([\u4e00-\u9fff（）\(\)，、\s]+)\s*([A-Z]\d{2,3}(?:\.\d{1,2})?)$");
                if (diagIcdMatch.Success) {
                    diagCode = diagIcdMatch.Groups[2].Value;
                    diagText = diagIcdMatch.Groups[1].Value;
                    source = beforeIcd.Substring(0, diagIcdMatch.Index).Trim();
                }

                // Everything after ICD code = ATC3 name + ingredient + drug name + usage + numbers
                var afterIcd = rowText.Substring(icdMatch.Index + icdMatch.Length);

                // Parse tail: find frequency pattern + numbers from the end
                var parsed = ParseTail(afterIcd);
                if (parsed != null) {
                    drugName = parsed[0];
                    usage = parsed[1];
                    days = parsed[2];
                    if (parsed.Length > 3) ingredientName = parsed[3];
                    if (parsed.Length > 4) totalQty = parsed[4];
                }
            }

            if (drugName == "" || usage == "") continue;

            // Clean source
            source = Regex.Replace(source, @"\d{8,}", "").Trim(); // Remove long numbers (hospital codes)
            source = source.Replace("\n", "").Replace("\r", "");

            var gk = date + "_" + source + "_" + diagCode;
            var mk = date + "_" + drugName;
            if (!groups.ContainsKey(gk))
                groups[gk] = new MedGroup { Date = date, Source = source, Diagnosis = diagText, DiagCode = diagCode };
            if (!seen.Contains(mk)) {
                seen.Add(mk);
                groups[gk].Medicines.Add(new MedItem { Name = drugName, Ingredient = ingredientName, Usage = usage, Days = days, TotalQty = totalQty });
            }
        }

        return FormatGroups(groups.Values.OrderByDescending(g => g.Date).ToList());
    }

    // Parse the tail of a row: extract drug name, usage, days from the combined string
    // Input: "ATC3名稱...成分名稱DRUG NAME 10MGQDPO30300Y"
    // Returns: [drugName, usage, days] or null
    static string[] ParseTail(string s) {
        if (string.IsNullOrEmpty(s)) return null;

        // Try to find frequency pattern near the end
        // Search from the right for known frequency codes
        var freqMatch = Regex.Match(s, @"(" + FreqPattern + @")(" + RoutePattern + @")?(\d+)(Y?)$");
        if (!freqMatch.Success) return null;

        var freq = freqMatch.Groups[1].Value;
        var numStr = freqMatch.Groups[3].Value;
        // Everything before the frequency is ATC3 + ingredient + drug name
        var beforeFreq = s.Substring(0, freqMatch.Index);

        // Extract drug name and ingredient
        string drugName, ingredient;
        ExtractDrugAndIngredient(beforeFreq, out ingredient, out drugName);

        // Parse numbers: extract totalQty and days
        string totalQty, days;
        ParseNumbers(numStr, out totalQty, out days);

        return new[] { drugName, freq, days, ingredient, totalQty };
    }

    // Parse concatenated number string into (totalQty, days)
    // e.g., "30300" → qty=30, days=30;  "130" → qty=1, days=30;  "220" → qty=2, days=2
    static void ParseNumbers(string numStr, out string totalQty, out string days) {
        totalQty = "0"; days = "0";
        if (string.IsNullOrEmpty(numStr)) return;

        var commonDays = new[] { 90, 60, 30, 28, 14, 7, 3, 2, 1 };
        foreach (var cd in commonDays) {
            var cdStr = cd.ToString();
            for (int pos = 1; pos <= numStr.Length - cdStr.Length; pos++) {
                if (numStr.Substring(pos, cdStr.Length) == cdStr) {
                    totalQty = numStr.Substring(0, pos);
                    days = cdStr;
                    return;
                }
            }
        }
        if (numStr.Length >= 2) {
            int half = numStr.Length / 2;
            totalQty = numStr.Substring(0, half);
            days = numStr.Substring(half);
        } else {
            days = numStr;
        }
    }

    static void ExtractDrugAndIngredient(string text, out string ingredient, out string drugName) {
        ingredient = "";
        drugName = "";
        if (string.IsNullOrEmpty(text)) return;

        // Step 1: Strip ATC3 name (Chinese or English) from the beginning
        // Patterns: "鈣通道阻滯劑(Calcium channel blockers)Amlodipine..."
        //           "Psycholeptics Zolpidem..."
        //           "口腔病藥物（Stomatological preparations）Triamcinolone..."
        string englishPart = text;

        // A) Find last Chinese char or full-width paren → strip everything up to it
        int lastCnPos = -1;
        for (int i = text.Length - 1; i >= 0; i--) {
            char c = text[i];
            if (c >= '\u4e00' && c <= '\u9fff' || c == '）') { lastCnPos = i; break; }
        }
        if (lastCnPos >= 0 && lastCnPos < text.Length - 3) {
            englishPart = text.Substring(lastCnPos + 1).Trim();
        }
        // B) If still starts with "(English ATC name)", strip it
        var atcParen = Regex.Match(englishPart, @"^\([^)]+\)\s*");
        if (atcParen.Success) {
            englishPart = englishPart.Substring(atcParen.Length);
        }
        // C) If starts with bare English ATC3 class name (long word, >=10 chars)
        // e.g., "PsycholepticsZolpidem" — ATC3 names are long (Psycholeptics=14)
        // Short words like "Aspirin"(7) are ingredient names, not ATC3
        var bareAtc = Regex.Match(englishPart, @"^([A-Z][a-z]{9,})(?=[A-Z])");
        if (bareAtc.Success) {
            englishPart = englishPart.Substring(bareAtc.Length);
        }

        // Step 2: englishPart = "IngredientNameDRUG NAME" or "IngredientNameDrug Name 0.2mg"
        // Find the boundary between ingredient and drug brand name
        // Approach: look for pattern where a lowercase letter is immediately followed
        // by an uppercase letter that starts a brand name (e.g., "TartrateSTILNOX")
        // Or Title Case followed by ALL CAPS

        // Strip leading "Y" (複方 flag from NHI cloud)
        if (englishPart.Length > 1 && englishPart[0] == 'Y' && char.IsUpper(englishPart[1])) {
            englishPart = englishPart.Substring(1);
        }

        // Key insight: ingredient names and drug names are concatenated without delimiter
        // e.g., "Tamsulosin HclHarnalidge D tablets 0.2mg"
        // e.g., "Pitavastatin CalciumLIVALO Tablets 2mg"
        // e.g., "AspirinBOKEY ENTERIC-MICROENCAPSULATED CAPSULES 100MG(ASPIRIN)"
        // e.g., "Sodium ChlorideSALINE INJECTION 0.9%"
        //
        // Strategy: scan for the boundary where lowercase→uppercase with no space
        // This is where ingredient ends and brand name starts

        // Find ALL positions where a lowercase letter is directly followed by an uppercase letter
        var candidates = new List<int>();
        for (int i = 1; i < englishPart.Length; i++) {
            if (char.IsLower(englishPart[i-1]) && char.IsUpper(englishPart[i])) {
                candidates.Add(i);
            }
        }

        // The LAST such boundary is most likely the ingredient→drug split
        if (candidates.Count > 0) {
            var splitPos = candidates[candidates.Count - 1];
            var drugPart = englishPart.Substring(splitPos);
            if (drugPart.Length >= 3) {
                ingredient = CleanIngredient(englishPart.Substring(0, splitPos));
                drugName = drugPart;
                return;
            }
        }

        // No camelCase boundary found — try finding ALL CAPS segment
        var brandSplit = Regex.Match(englishPart, @"([A-Z]{3}[A-Z\d\s\.\-\(\)\""\'%/,]*)$");
        if (brandSplit.Success) {
            ingredient = CleanIngredient(englishPart.Substring(0, brandSplit.Index));
            drugName = brandSplit.Groups[1].Value.Trim();
            return;
        }

        // Fallback
        drugName = englishPart;
    }

    // Clean ingredient name: remove salt forms, keep core name
    static string CleanIngredient(string raw) {
        var s = raw.Trim();
        // Remove compound separator — take only first ingredient
        if (s.Contains("；")) s = s.Split('；')[0].Trim();
        if (s.Contains(";")) s = s.Split(';')[0].Trim();
        // Remove parenthetical salt/ester forms
        s = Regex.Replace(s, @"\s*\((?:Besylate|Sodium|Calcium|Tartrate|Acetonide|Hcl|Hydrochloride|Maleate|Fumarate|Mesylate|Potassium|As\s+\w+)\)", "", RegexOptions.IgnoreCase);
        // Remove trailing salt words
        s = Regex.Replace(s, @"\s+(?:Sodium|Calcium|Tartrate|Hcl|Hydrochloride|Besylate|Maleate|Fumarate|Mesylate|Potassium|17-Propionate)\s*$", "", RegexOptions.IgnoreCase);
        return s.Trim();
    }

    // ── Format output: Ingredient(dose) usage, horizontal ──
    static string FormatGroups(List<MedGroup> groups) {
        if (groups == null || groups.Count == 0) return null;
        // Flatten all medicines, skip STAT (emergency IV/injections)
        var medParts = new List<string>();
        foreach (var g in groups) {
            foreach (var m in g.Medicines) {
                if (m.Usage == "STAT") continue;
                var displayName = FormatMedDisplay(m);
                var sep = displayName.EndsWith(")") ? "" : " ";
                medParts.Add(displayName + sep + CleanUsage(m.Usage));
            }
        }
        return string.Join(", ", medParts);
    }

    // Strip PC/AC/HS suffixes: QDPC→QD, BIDPC→BID, HSPC→HS
    static string CleanUsage(string u) {
        if (string.IsNullOrEmpty(u)) return u;
        return Regex.Replace(u, @"(QD|BID|TID|QID|HS)(PC|AC|EXT|PO)$", "$1");
    }

    // Format a single medicine: Ingredient(dose) with optional N# for multi-tablet
    static string FormatMedDisplay(MedItem m) {
        // Extract dose from drug name (e.g., "NORVASC TABLETS 5MG" → "5")
        string dose = "";
        var doseMatch = Regex.Match(m.Name, @"(\d+(?:\.\d+)?)\s*(?:mg|mcg|g|ml|%)", RegexOptions.IgnoreCase);
        if (doseMatch.Success) dose = doseMatch.Groups[1].Value;

        // Prefer ingredient name if available
        string name = "";
        if (!string.IsNullOrEmpty(m.Ingredient) && m.Ingredient.Length >= 2) {
            name = m.Ingredient;
        } else {
            name = SimplifyName(m.Name);
            name = Regex.Replace(name, @"\s*\(\d+(?:\.\d+)?\)\s*", " ").Trim();
        }

        // Calculate per-dose quantity
        string perDose = CalcPerDose(m.TotalQty, m.Usage, m.Days);

        var sb = new StringBuilder();
        sb.Append(name);
        if (dose != "") sb.Append("(").Append(dose).Append(")");
        // Show N# only when > 1 tablet per dose
        if (perDose != "" && perDose != "1") sb.Append(" ").Append(perDose).Append("#");
        return sb.ToString();
    }

    // Calculate per-dose tablet count from total quantity, usage, days
    static string CalcPerDose(string totalQty, string usage, string days) {
        if (string.IsNullOrEmpty(totalQty) || string.IsNullOrEmpty(usage) || string.IsNullOrEmpty(days))
            return "";
        double total; int numDays;
        if (!double.TryParse(totalQty, out total) || !int.TryParse(days, out numDays) || numDays <= 0 || total <= 0)
            return "";

        var freq = (usage ?? "").ToUpper();
        int mult = 1;
        if (freq.Contains("BID")) mult = 2;
        else if (freq.Contains("TID")) mult = 3;
        else if (freq.Contains("QID")) mult = 4;
        else if (freq.Contains("Q6H")) mult = 4;
        else if (freq.Contains("Q8H")) mult = 3;
        else if (freq.Contains("Q12H")) mult = 2;
        else if (freq.Contains("Q4H")) mult = 6;
        else if (freq.Contains("Q2H")) mult = 12;
        // QD, HS, QN, etc. = 1

        int totalDoses = numDays * mult;
        if (totalDoses <= 0) return "";
        double single = total / totalDoses;
        double rounded = Math.Round(single * 4) / 4.0;
        if (rounded < 0.5) return "";
        if (rounded == 1.0) return "1";
        return rounded == (int)rounded ? ((int)rounded).ToString() : rounded.ToString("0.##");
    }

    public static string SimplifyName(string name) {
        var s = name;
        s = RxTablet.Replace(s, "");
        s = RxCapsule.Replace(s, "");
        s = RxDoseMg.Replace(s, m => "(" + m.Groups[1].Value.Replace(" ", "") + ")");
        s = Regex.Replace(s, @"\([^)]*箔[^)]*\)", "");  // Remove packaging info
        s = Regex.Replace(s, @"""[A-Za-z\.\s]+""", "");  // Remove quoted manufacturer names like "SINPHAR"
        s = s.Replace("\"", "");
        s = Regex.Replace(s, @"\s+", " ").Trim();
        var dm = RxComplexDose.Match(s);
        if (dm.Success) {
            var dose = dm.Groups[1].Value.Replace(" ", "");
            s = s.Remove(dm.Index, dm.Length).Trim() + " (" + dose + ")";
        }
        return s;
    }
}

// ══════════ Config ══════════
class Slot {
    public string Name="",Type="none",Text="",Hotkey="";
    public List<string> LabItems=new List<string>();
}
static class VK {
    public static uint ModFlag(string mod) { return mod=="Alt"?(uint)1:(uint)2; }
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
        return 0xC0;
    }
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
    public string CaptureKey = "`";
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
    static string Path { get { if(_path==null) _path=System.IO.Path.Combine(ConfigDir,"lab_formatter_config.json"); return _path; } }

    public void Load() {
        Slots.Clear();
        try {
            if (!File.Exists(Path)) { CreateDefault(); }
            var d = Json.Parse(File.ReadAllText(Path, Encoding.UTF8));
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
            if(i==2) Slots.Add(new Slot{Name="濃縮報告",Type="lab",
                Hotkey="Ctrl+"+(i+1),
                LabItems=Lab.Items.Where(x=>x.On).Select(x=>x.Key).ToList()});
            else Slots.Add(new Slot{Name="快貼"+(i<2?i+1:i),Type="paste",
                Hotkey="Ctrl+"+(i+1)});
        }
    }
    public void Save() {
        var doc = new Dictionary<string,object>();
        doc["_說明"]="Lab Data Formatter 設定檔";
        var fmt=new Dictionary<string,object>();
        fmt["type"]=new List<object>{"paste=快貼文字","lab=報告整理","template=範本套用","none=未啟用"};
        fmt["熱鍵"]="Ctrl+0=設定, Ctrl+1~4=slots[0]~[3]";
        doc["_格式說明"]=fmt;
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
        File.WriteAllText(Path, Json.Encode(doc,0), Encoding.UTF8);
    }
    void CreateDefault() {
        var defaults = new List<Slot>();
        defaults.Add(new Slot{Name="快貼1",Type="paste",Hotkey="Ctrl+1"});
        defaults.Add(new Slot{Name="快貼2",Type="paste",Hotkey="Ctrl+2"});
        defaults.Add(new Slot{Name="濃縮報告",Type="lab",Hotkey="Ctrl+3",
            LabItems=Lab.Items.Where(x=>x.On).Select(x=>x.Key).ToList()});
        defaults.Add(new Slot{Name="快貼3",Type="paste",Hotkey="Ctrl+4"});
        var doc = new Dictionary<string,object>();
        doc["_說明"]="Lab Data Formatter 設定檔";
        var fmt=new Dictionary<string,object>();
        fmt["type"]=new List<object>{"paste=快貼文字","lab=報告整理","template=範本套用","none=未啟用"};
        fmt["熱鍵"]="Ctrl+0=設定, Ctrl+1~4=slots[0]~[3]";
        doc["_格式說明"]=fmt;
        var arr = new List<object>();
        for(int i=0;i<defaults.Count;i++){var s=defaults[i];
            var d=new Dictionary<string,object>();
            d["_hotkey"]="Ctrl+"+(i+1); d["name"]=s.Name; d["type"]=s.Type;
            if(s.Type=="paste"||s.Type=="template") d["text"]=s.Text;
            if(s.Type=="lab") d["lab_items"]=s.LabItems.Cast<object>().ToList();
            arr.Add(d);}
        doc["slots"]=arr;
        doc["capture_mod"]="Alt"; doc["capture_key"]="1";
        doc["paste_mod"]="Alt"; doc["paste_key"]="2";
        try{File.WriteAllText(Path, Json.Encode(doc,0), Encoding.UTF8);}catch{}
    }
    public HashSet<string> GetLabItems() {
        var s=Slots.FirstOrDefault(x=>x.Type=="lab");
        if(s!=null&&s.LabItems.Count>0) return new HashSet<string>(s.LabItems);
        return new HashSet<string>(Lab.Items.Where(x=>x.On).Select(x=>x.Key));
    }
    public static void OpenJson() { try{System.Diagnostics.Process.Start(Path);}catch{} }
    public static void OpenFolder() { try{System.Diagnostics.Process.Start("explorer.exe",ConfigDir);}catch{} }
    public static string JsonPath { get { return Path; } }
}

// ══════════ Native helpers ══════════
static class NativeMethods {
    [DllImport("user32")] public static extern bool ReleaseCapture();
    [DllImport("user32")] public static extern int SendMessage(IntPtr h, int msg, int wp, int lp);
    // DWM for Mica/Acrylic backdrop
    [DllImport("dwmapi.dll")] public static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int val, int size);
    public const int DWMWA_USE_IMMERSIVE_DARK_MODE = 20;
    public const int DWMWA_SYSTEMBACKDROP_TYPE = 38;
}

// ══════════ Main App ══════════
[StructLayout(LayoutKind.Sequential)]
struct KEYBDINPUT { public ushort wVk; public ushort wScan; public uint dwFlags; public uint time; public IntPtr dwExtraInfo; }
[StructLayout(LayoutKind.Sequential)]
struct MOUSEINPUT { public int dx,dy; public uint mouseData,dwFlags,time; public IntPtr dwExtraInfo; }
[StructLayout(LayoutKind.Explicit)]
struct INPUTUNION { [FieldOffset(0)] public MOUSEINPUT mi; [FieldOffset(0)] public KEYBDINPUT ki; }
[StructLayout(LayoutKind.Sequential)]
struct INPUT { public uint type; public INPUTUNION u; }

class App : Form {
    const string VER="v2.0";
    [DllImport("user32")] static extern bool RegisterHotKey(IntPtr h,int id,uint mod,uint vk);
    [DllImport("user32")] static extern bool UnregisterHotKey(IntPtr h,int id);
    [DllImport("user32")] static extern bool DestroyIcon(IntPtr handle);
    [DllImport("user32")] static extern uint SendInput(uint n, INPUT[] inputs, int size);
    [DllImport("user32")] static extern bool AddClipboardFormatListener(IntPtr hwnd);
    [DllImport("user32")] static extern bool RemoveClipboardFormatListener(IntPtr hwnd);
    const int WM_CLIPBOARDUPDATE = 0x031D;

    NotifyIcon tray; Config cfg = new Config();
    string labResult; string medResult; string lastClip; DateTime ignoreUntil;
    const uint MOD_CTRL=2; const uint MOD_ALT=1; const int WM_HOTKEY=0x312;
    static readonly Dictionary<int,uint> HK = new Dictionary<int,uint>{
        {0,0x30},{1,0x31},{2,0x32},{3,0x33},{4,0x34}};
    const int HKID_CAPTURE=5;
    const int HKID_PASTE=6;

    // Buffers
    List<string> rawBuffer = new List<string>();         // ORD lab
    List<string> cloudLabBuffer = new List<string>();     // Cloud lab
    List<string> cloudMedBuffer = new List<string>();     // Cloud medication
    System.Windows.Forms.Timer mergeTimer = new System.Windows.Forms.Timer{Interval=15000};

    Icon MakeIcon(Color c) {
        var bmp=new Bitmap(32,32); using(var g=Graphics.FromImage(bmp)){
            g.SmoothingMode=SmoothingMode.AntiAlias;
            g.Clear(Color.Transparent);
            using(var b=new SolidBrush(Color.FromArgb(35,c)))g.FillEllipse(b,0,3,31,28);
            using(var b=new SolidBrush(Color.FromArgb(255,225,195)))g.FillEllipse(b,3,10,26,21);
            using(var b=new SolidBrush(Color.FromArgb(90,55,30))){
                g.FillEllipse(b,3,11,7,8); g.FillEllipse(b,22,11,7,8);}
            using(var b=new SolidBrush(Color.White)){
                g.FillEllipse(b,5,0,22,14); g.FillRectangle(b,4,7,24,6);}
            using(var b=new SolidBrush(c)){
                g.FillRectangle(b,14,2,4,8); g.FillRectangle(b,11,4,10,4);}
            using(var p=new Pen(Color.FromArgb(180,200,200,200),0.8f))g.DrawLine(p,5,12,27,12);
            using(var b=new SolidBrush(Color.FromArgb(45,45,55))){
                g.FillEllipse(b,8,16,6,6); g.FillEllipse(b,18,16,6,6);}
            using(var b=new SolidBrush(Color.White)){
                g.FillEllipse(b,9,16,3,3); g.FillEllipse(b,19,16,3,3);}
            using(var b=new SolidBrush(Color.FromArgb(220,255,255,255))){
                g.FillEllipse(b,12,20,2,1); g.FillEllipse(b,22,20,2,1);}
            using(var b=new SolidBrush(Color.FromArgb(65,255,100,110))){
                g.FillEllipse(b,4,22,8,4); g.FillEllipse(b,20,22,8,4);}
            using(var p=new Pen(Color.FromArgb(210,90,70),1.2f)){
                g.DrawArc(p,10,23,6,3,10,160); g.DrawArc(p,16,23,6,3,10,160);}
        }
        var hIcon=bmp.GetHicon(); var icon=Icon.FromHandle(hIcon);
        var clone=(Icon)icon.Clone(); DestroyIcon(hIcon); bmp.Dispose();
        return clone;
    }

    public App() {
        cfg.Load(); ShowInTaskbar=false; WindowState=FormWindowState.Minimized; Visible=false;
        tray=new NotifyIcon{Icon=MakeIcon(Theme.Gold),Text="Lab Formatter "+VER,Visible=true};
        var menu=new ContextMenuStrip();
        menu.Items.Add("Ctrl+0 設定",null,(s,e)=>ShowSettings());
        menu.Items.Add("編輯 JSON",null,(s,e)=>Config.OpenJson());
        menu.Items.Add("手動轉換",null,(s,e)=>ShowManual());
        menu.Items.Add("使用說明",null,(s,e)=>OpenReadme());
        menu.Items.Add(new ToolStripSeparator());
        menu.Items.Add("結束",null,(s,e)=>{tray.Visible=false;Application.Exit();});
        tray.ContextMenuStrip=menu;
        tray.DoubleClick+=(s,e)=>ShowSettings();

        foreach(var kv in HK) RegisterHotKey(Handle,kv.Key,MOD_CTRL,kv.Value);
        RegisterCaptureHotkey();
        RegisterPasteHotkey();

        lastClip=SafeGetClip(); ignoreUntil=DateTime.MinValue;
        mergeTimer.Tick+=(s,e)=>{mergeTimer.Stop(); FinalizeMerge();};
        previewTimer.Tick+=(s,e)=>{previewTimer.Stop();if(previewForm!=null&&!previewForm.IsDisposed)previewForm.Hide();};
        yellowTimer.Tick+=(s,e)=>{yellowTimer.Stop();
            SetTray(Theme.Gold,labResult!=null||medResult!=null?
                cfg.PasteMod+"+"+cfg.PasteKey+" 貼上 | "+(labResult!=null?labResult.Substring(0,Math.Min(30,labResult.Length)):"藥歷"):
                cfg.CaptureMod+"+"+cfg.CaptureKey+" 擷取");};
        AddClipboardFormatListener(Handle);
    }

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

    // ── Clipboard change handler (expanded for 3 data types) ──
    void OnClipboardChanged() {
        var t=SafeGetClip();
        if(t==""||t==lastClip) return;
        lastClip=t;
        if(DateTime.Now<ignoreUntil) return;

        // Priority: ORD lab → Cloud lab → Cloud medication
        if (Lab.IsLabData(t)) {
            rawBuffer.Add(t);
            var combined=string.Join("\n",rawBuffer);
            var r=Lab.Convert(combined,cfg.GetLabItems());
            if(r!=null){
                labResult=r;
                var countTag=rawBuffer.Count>1?" ["+rawBuffer.Count+"筆]":"";
                SetTray(Theme.Accent,"✓"+countTag+" "+r.Substring(0,Math.Min(50,r.Length)));
                ShowPreview(r, null, rawBuffer.Count);
                ResetMergeTimer();
            }
        }
        else if (CloudLab.IsCloudLabData(t)) {
            cloudLabBuffer.Add(t);
            var combined=string.Join("\n",cloudLabBuffer);
            var r=CloudLab.Convert(combined,cfg.GetLabItems());
            if(r!=null){
                labResult=r;
                var countTag=cloudLabBuffer.Count>1?" [雲端"+cloudLabBuffer.Count+"筆]":"[雲端]";
                SetTray(Theme.Accent,"✓"+countTag+" "+r.Substring(0,Math.Min(50,r.Length)));
                ShowPreview(r, null, cloudLabBuffer.Count);
                ResetMergeTimer();
            }
        }
        else if (CloudMed.IsCloudMedData(t)) {
            cloudMedBuffer.Add(t);
            var combined=string.Join("\n",cloudMedBuffer);
            var r=CloudMed.Convert(combined);
            if(r!=null){
                medResult=r;
                SetTray(Theme.Blue,"✓ [藥歷] "+r.Split('\n')[0]);
                ShowPreview(labResult, r, 0);
                ResetMergeTimer();
            }
        }
    }

    void ResetMergeTimer() { mergeTimer.Stop(); mergeTimer.Start(); }
    void FinalizeMerge() {
        rawBuffer.Clear(); cloudLabBuffer.Clear(); cloudMedBuffer.Clear();
        DelayYellow();
    }

    string GetPasteResult() {
        var sb = new StringBuilder();
        if (labResult != null) sb.Append(labResult);
        if (medResult != null) {
            if (sb.Length > 0) sb.AppendLine().AppendLine();
            sb.Append(medResult);
        }
        return sb.Length > 0 ? sb.ToString() : null;
    }

    // ══════ Apple-style Floating Preview ══════
    Form previewForm;
    Label previewLabel;
    Label previewTitle;
    System.Windows.Forms.Timer previewTimer = new System.Windows.Forms.Timer();

    void ShowPreview(string labText, string medText, int labCount) {
        if(previewForm==null||previewForm.IsDisposed) {
            previewForm=new Form{
                FormBorderStyle=FormBorderStyle.None,
                ShowInTaskbar=false,
                TopMost=true,
                BackColor=Theme.Bg,
                Opacity=0.94,
                StartPosition=FormStartPosition.Manual,
                Size=new Size(520,80)
            };
            // Try Mica backdrop (Win11)
            try {
                int val = 1;
                NativeMethods.DwmSetWindowAttribute(previewForm.Handle, NativeMethods.DWMWA_USE_IMMERSIVE_DARK_MODE, ref val, sizeof(int));
                val = 2; // Mica
                NativeMethods.DwmSetWindowAttribute(previewForm.Handle, NativeMethods.DWMWA_SYSTEMBACKDROP_TYPE, ref val, sizeof(int));
            } catch { }

            var topBar=new Panel{Dock=DockStyle.Top,Height=28,BackColor=Theme.Surface};
            topBar.Paint+=(s,e)=>{
                using(var b=new SolidBrush(Theme.Surface))
                    e.Graphics.FillRectangle(b,0,0,topBar.Width,topBar.Height);
            };
            previewTitle=new Label{
                ForeColor=Theme.Accent,Font=new Font("Microsoft JhengHei UI",9,FontStyle.Bold),
                Dock=DockStyle.Fill,TextAlign=ContentAlignment.MiddleLeft,Padding=new Padding(12,0,0,0)};
            var closeBtn=new Label{Text="✕",ForeColor=Theme.Text2,
                Font=new Font("Arial",9),Dock=DockStyle.Right,Width=32,
                TextAlign=ContentAlignment.MiddleCenter,Cursor=Cursors.Hand};
            closeBtn.Click+=(s,e)=>previewForm.Hide();
            closeBtn.MouseEnter+=(s,e)=>closeBtn.ForeColor=Theme.Red;
            closeBtn.MouseLeave+=(s,e)=>closeBtn.ForeColor=Theme.Text2;
            topBar.Controls.Add(previewTitle); topBar.Controls.Add(closeBtn);
            previewLabel=new Label{ForeColor=Theme.Text,
                Font=Theme.Mono,Dock=DockStyle.Fill,
                TextAlign=ContentAlignment.TopLeft,Padding=new Padding(12,8,12,8)};
            previewForm.Controls.Add(previewLabel);
            previewForm.Controls.Add(topBar);
            topBar.MouseDown+=(s,e)=>{if(e.Button==MouseButtons.Left){
                NativeMethods.ReleaseCapture();NativeMethods.SendMessage(previewForm.Handle,0xA1,0x2,0);}};
            previewTitle.MouseDown+=(s,e)=>{if(e.Button==MouseButtons.Left){
                NativeMethods.ReleaseCapture();NativeMethods.SendMessage(previewForm.Handle,0xA1,0x2,0);}};
        }

        // Build display text
        var display = new StringBuilder();
        if (labText != null) display.Append(labText);
        if (medText != null) {
            if (display.Length > 0) display.AppendLine().AppendLine("── 藥歷 ──");
            // Show first few lines of medication
            var medLines = medText.Split('\n');
            int maxLines = Math.Min(medLines.Length, 8);
            for (int i = 0; i < maxLines; i++) display.AppendLine(medLines[i]);
            if (medLines.Length > maxLines) display.Append("  ... 共 " + medLines.Length + " 行");
        }

        // Title
        var titleParts = new List<string>();
        if (labText != null) {
            if (labCount > 1) titleParts.Add("✓ " + labCount + "筆報告");
            else titleParts.Add("✓ 報告已轉換");
        }
        if (medText != null) titleParts.Add("✓ 藥歷已整理");
        titleParts.Add(cfg.PasteMod + "+" + cfg.PasteKey + " 貼上");
        previewTitle.Text = string.Join(" | ", titleParts);

        previewLabel.Text = display.ToString();

        // Size
        var lines = display.ToString().Split('\n');
        int maxLen=0; foreach(var ln in lines) if(ln.Length>maxLen) maxLen=ln.Length;
        int textW=(int)(maxLen*7.5)+40;
        if(textW<420) textW=420;
        if(textW>Screen.PrimaryScreen.WorkingArea.Width-20) textW=Screen.PrimaryScreen.WorkingArea.Width-20;
        int textH=28+Math.Max(lines.Length,1)*18+20;
        if(textH<70) textH=70;
        if(textH>400) textH=400;
        previewForm.Size=new Size(textW,textH);
        Theme.SetRoundRegion(previewForm, 12);
        var wa=Screen.PrimaryScreen.WorkingArea;
        previewForm.Location=new Point(wa.Right-previewForm.Width-12, wa.Bottom-previewForm.Height-12);
        previewForm.Show();
        int hideDelay = (labText != null && medText != null) ? 18000 : labCount > 1 ? 15000 : 8000;
        previewTimer.Interval=hideDelay;
        previewTimer.Stop(); previewTimer.Start();
    }
    void HidePreview(){
        if(previewForm!=null&&!previewForm.IsDisposed&&previewForm.Visible) previewForm.Hide();
        previewTimer.Stop();
    }
    void SetTray(Color c,string tip){
        var old=tray.Icon; tray.Icon=MakeIcon(c); if(old!=null)old.Dispose();
        tray.Text=tip.Length>63?tip.Substring(0,63):tip;}
    System.Windows.Forms.Timer yellowTimer = new System.Windows.Forms.Timer{Interval=3000};
    void DelayYellow(){ yellowTimer.Stop(); yellowTimer.Start(); }
    void WriteClip(string t){ignoreUntil=DateTime.Now.AddMilliseconds(800);
        SafeSetClip(t); lastClip=t;}

    static void SendKeys(params ushort[] vks){
        var inputs=new List<INPUT>();
        foreach(var vk in vks){
            var inp=new INPUT(); inp.type=1; inp.u.ki.wVk=vk; inp.u.ki.dwFlags=0;
            inputs.Add(inp);
        }
        for(int i=vks.Length-1;i>=0;i--){
            var inp=new INPUT(); inp.type=1; inp.u.ki.wVk=vks[i]; inp.u.ki.dwFlags=2;
            inputs.Add(inp);
        }
        SendInput((uint)inputs.Count, inputs.ToArray(), Marshal.SizeOf(typeof(INPUT)));
    }
    static void ReleaseAllModifiers(){
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
            SendKeys(0x11, 0x56);
        });
    }
    void SimCtrlAC(){
        ThreadPool.QueueUserWorkItem(delegate{
            Thread.Sleep(50);
            ReleaseAllModifiers(); Thread.Sleep(30);
            SendKeys(0x11, 0x41);
            Thread.Sleep(80);
            SendKeys(0x11, 0x43);
        });
    }

    bool IsCaptureAllowed(){ return true; }
    void RegisterCaptureHotkey(){
        UnregisterHotKey(Handle,HKID_CAPTURE);
        uint mod=VK.ModFlag(cfg.CaptureMod);
        uint vk=VK.Code(cfg.CaptureKey);
        if(!RegisterHotKey(Handle,HKID_CAPTURE,mod,vk))
            tray.ShowBalloonTip(3000,"擷取鍵註冊失敗",
                cfg.CaptureMod+"+"+cfg.CaptureKey+" 被其他程式佔用，請在設定中更換熱鍵",ToolTipIcon.Warning);
    }
    void RegisterPasteHotkey(){
        UnregisterHotKey(Handle,HKID_PASTE);
        uint mod=VK.ModFlag(cfg.PasteMod);
        uint vk=VK.Code(cfg.PasteKey);
        if(!RegisterHotKey(Handle,HKID_PASTE,mod,vk))
            tray.ShowBalloonTip(3000,"貼上鍵註冊失敗",
                cfg.PasteMod+"+"+cfg.PasteKey+" 被其他程式佔用，請在設定中更換熱鍵",ToolTipIcon.Warning);
    }

    void DoPaste() {
        // Finalize any pending buffer
        if(rawBuffer.Count>0){
            labResult=Lab.Convert(string.Join("\n",rawBuffer),cfg.GetLabItems());
            rawBuffer.Clear(); if(mergeTimer!=null) mergeTimer.Stop();
        }
        if(cloudLabBuffer.Count>0){
            labResult=CloudLab.Convert(string.Join("\n",cloudLabBuffer),cfg.GetLabItems());
            cloudLabBuffer.Clear(); if(mergeTimer!=null) mergeTimer.Stop();
        }
        if(cloudMedBuffer.Count>0){
            medResult=CloudMed.Convert(string.Join("\n",cloudMedBuffer));
            cloudMedBuffer.Clear(); if(mergeTimer!=null) mergeTimer.Stop();
        }
        var result = GetPasteResult();
        if(result!=null){WriteClip(result);SimCtrlV();
            HidePreview();SetTray(Theme.Accent,"✓ 已貼上");DelayYellow();}
        else{SetTray(Theme.Red,"請先擷取報告");DelayYellow();}
    }

    protected override void WndProc(ref Message m) {
        if(m.Msg==WM_CLIPBOARDUPDATE){ OnClipboardChanged(); base.WndProc(ref m); return; }
        if(m.Msg==WM_HOTKEY){int id=(int)m.WParam;
            if(id==0){ShowSettings();return;}
            if(id==HKID_CAPTURE){
                if(IsCaptureAllowed()) SimCtrlAC();
                return;
            }
            if(id==HKID_PASTE){
                DoPaste();
                return;
            }
            int idx=id-1; if(idx<0||idx>=cfg.Slots.Count)return;
            var s=cfg.Slots[idx];
            try{
                if(s.Type=="paste"){if(s.Text!=""){WriteClip(s.Text);SimCtrlV();}}
                else if(s.Type=="lab"){ DoPaste(); }
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

    // ══════ Apple-style Settings Window ══════
    Form settingsForm;
    void ShowSettings() {
        if(settingsForm!=null&&!settingsForm.IsDisposed){settingsForm.BringToFront();return;}
        cfg.Load();
        var f=new Form{Text="Lab Formatter "+VER+" 設定",Size=new Size(640,720),
            StartPosition=FormStartPosition.CenterScreen,TopMost=true,
            MaximizeBox=false,BackColor=Theme.Bg,ForeColor=Theme.Text,
            Font=Theme.Body};
        settingsForm=f;
        // Dark title bar
        try{int val=1;NativeMethods.DwmSetWindowAttribute(f.Handle,NativeMethods.DWMWA_USE_IMMERSIVE_DARK_MODE,ref val,sizeof(int));}catch{}

        int fw=f.ClientSize.Width;
        var scroll=new Panel{AutoScroll=true,Dock=DockStyle.Fill};
        f.Controls.Add(scroll);
        int yy=8;

        // ── Header ──
        var header=new Label{Text=cfg.CaptureMod+"+"+cfg.CaptureKey+" 擷取 | "+cfg.PasteMod+"+"+cfg.PasteKey+" 貼上 | Ctrl+0 設定",
            Left=0,Top=yy,Width=fw,Height=24,TextAlign=ContentAlignment.MiddleCenter,ForeColor=Theme.Text3};
        scroll.Controls.Add(header); yy+=30;

        // ── Slot editors ──
        var slots=new InlineSlot[4];
        for(int i=0;i<4;i++){
            slots[i]=new InlineSlot(scroll,cfg.Slots[i],i,yy,fw-24);
            yy+=slots[i].Height+6;
        }

        // ── Item checkboxes card ──
        var itemCard=new CardPanel{Title="輸出項目",Left=8,Top=yy,Width=fw-16,Height=160};
        scroll.Controls.Add(itemCard); yy+=itemCard.Height+8;
        var chkItems=new Dictionary<string,CheckBox>();
        var labEnabled=cfg.GetLabItems();
        int cx=12,cy=38;
        foreach(var it in Lab.Items){
            var cb=new CheckBox{Text=it.Disp,Left=cx,Top=cy,Width=85,Height=20,
                Checked=labEnabled.Contains(it.Key),ForeColor=Theme.Text,
                Font=Theme.Small,BackColor=Color.Transparent};
            chkItems[it.Key]=cb; itemCard.Controls.Add(cb);
            cx+=90; if(cx>fw-120){cx=12;cy+=22;}
        }

        // ── Hotkey card ──
        var hkCard=new CardPanel{Title="熱鍵設定",Left=8,Top=yy,Width=fw-16,Height=90};
        scroll.Controls.Add(hkCard); yy+=hkCard.Height+8;
        // Capture
        hkCard.Controls.Add(new Label{Text="擷取鍵:",Left=12,Top=40,Width=55,ForeColor=Theme.Text2,BackColor=Color.Transparent});
        var capModCombo=new ComboBox{Left=68,Top=37,Width=60,DropDownStyle=ComboBoxStyle.DropDownList,
            BackColor=Theme.Surface,ForeColor=Theme.Text};
        capModCombo.Items.AddRange(new object[]{"Ctrl","Alt"}); capModCombo.SelectedItem=cfg.CaptureMod;
        hkCard.Controls.Add(capModCombo);
        hkCard.Controls.Add(new Label{Text="+",Left=130,Top=40,Width=12,ForeColor=Theme.Text2,BackColor=Color.Transparent});
        var capKeyBox=new TextBox{Left=144,Top=37,Width=40,Text=cfg.CaptureKey,MaxLength=5,
            BackColor=Theme.Surface,ForeColor=Theme.Accent,Font=Theme.MonoSmall,TextAlign=HorizontalAlignment.Center};
        hkCard.Controls.Add(capKeyBox);
        // Paste
        hkCard.Controls.Add(new Label{Text="貼上鍵:",Left=210,Top=40,Width=55,ForeColor=Theme.Text2,BackColor=Color.Transparent});
        var pasteModCombo=new ComboBox{Left=266,Top=37,Width=60,DropDownStyle=ComboBoxStyle.DropDownList,
            BackColor=Theme.Surface,ForeColor=Theme.Text};
        pasteModCombo.Items.AddRange(new object[]{"Ctrl","Alt"}); pasteModCombo.SelectedItem=cfg.PasteMod;
        hkCard.Controls.Add(pasteModCombo);
        hkCard.Controls.Add(new Label{Text="+",Left=328,Top=40,Width=12,ForeColor=Theme.Text2,BackColor=Color.Transparent});
        var pasteKeyBox=new TextBox{Left=342,Top=37,Width=40,Text=cfg.PasteKey,MaxLength=5,
            BackColor=Theme.Surface,ForeColor=Theme.Accent,Font=Theme.MonoSmall,TextAlign=HorizontalAlignment.Center};
        hkCard.Controls.Add(pasteKeyBox);

        // ── Data source info card ──
        var srcCard=new CardPanel{Title="支援的資料來源",Left=8,Top=yy,Width=fw-16,Height=70};
        scroll.Controls.Add(srcCard); yy+=srcCard.Height+8;
        srcCard.Controls.Add(new Label{Text="  ORD 檢驗報告 (Ctrl+A → Ctrl+C)    健保雲端檢驗    健保雲端藥歷",
            Left=12,Top=38,Width=fw-40,Height=20,ForeColor=Theme.Accent,BackColor=Color.Transparent,Font=Theme.Small});

        // ── Buttons ──
        var saveBtn=new PillButton{Text="儲存設定",Left=fw/2-130,Top=yy,Width=120,Height=34,IsPrimary=true};
        var helpBtn=new PillButton{Text="使用說明",Left=fw/2+10,Top=yy,Width=110,Height=34,IsPrimary=false,ForeColor=Theme.Text2};
        helpBtn.Click+=(s,e)=>OpenReadme();
        scroll.Controls.Add(saveBtn);scroll.Controls.Add(helpBtn);
        yy+=42;

        saveBtn.Click+=(s,e)=>{
            for(int i=0;i<4;i++) slots[i].SaveTo(cfg.Slots[i]);
            var checkedKeys=new List<string>();
            foreach(var kv in chkItems) if(kv.Value.Checked) checkedKeys.Add(kv.Key);
            foreach(var sl in cfg.Slots) if(sl.Type=="lab") sl.LabItems=checkedKeys;
            cfg.CaptureMod=capModCombo.Text;
            cfg.CaptureKey=capKeyBox.Text.Trim(); if(cfg.CaptureKey=="") cfg.CaptureKey="1";
            cfg.PasteMod=pasteModCombo.Text;
            cfg.PasteKey=pasteKeyBox.Text.Trim(); if(cfg.PasteKey=="") cfg.PasteKey="2";
            cfg.Save();
            RegisterCaptureHotkey(); RegisterPasteHotkey();
            SetTray(Theme.Accent,"✓ 已儲存");DelayYellow();
        };

        // JSON row
        var jlbl=new Label{Text="JSON:",Left=10,Top=yy+2,Width=38,ForeColor=Theme.Text3};
        var pe=new TextBox{Left=50,Top=yy,Width=fw-210,ReadOnly=true,Text=Config.JsonPath,
            BackColor=Theme.Surface,ForeColor=Theme.Text2,Font=new Font("Consolas",8),BorderStyle=BorderStyle.FixedSingle};
        var b1=new PillButton{Text="開啟",Left=fw-154,Top=yy-2,Width=55,Height=24,IsPrimary=false,ForeColor=Theme.Text2};
        b1.Click+=(s,e)=>Config.OpenJson();
        var b2=new PillButton{Text="讀取",Left=fw-94,Top=yy-2,Width=55,Height=24,IsPrimary=false,ForeColor=Theme.Text2};
        b2.Click+=(s,e)=>{cfg.Load();for(int i=0;i<4;i++)slots[i].LoadFrom(cfg.Slots[i]);
            SetTray(Theme.Accent,"✓ 已讀取");DelayYellow();};
        scroll.Controls.AddRange(new Control[]{jlbl,pe,b1,b2}); yy+=28;

        // Result
        scroll.Controls.Add(new Label{Text="最新:",Left=10,Top=yy+2,Width=40,ForeColor=Theme.Accent,
            Font=new Font("Microsoft JhengHei UI",9,FontStyle.Bold)});
        var st=new TextBox{Text=labResult!=null?labResult:"尚無結果",
            Left=52,Top=yy,Width=fw-62,Height=20,ReadOnly=true,BackColor=Theme.Surface,
            ForeColor=Theme.Accent,Font=Theme.MonoSmall,BorderStyle=BorderStyle.FixedSingle};
        scroll.Controls.Add(st); yy+=28;

        // Footer
        var feedback=new Label{Text="意見回饋: 吳岳霖醫師  DAL93@tpech.gov.tw    "+VER,
            Left=0,Top=yy,Width=fw,Height=18,TextAlign=ContentAlignment.MiddleRight,
            ForeColor=Theme.Text3,Font=Theme.Small,Padding=new Padding(0,0,10,0)};
        scroll.Controls.Add(feedback); yy+=24;

        f.ClientSize=new Size(fw,Math.Min(yy+8,720));
        f.ShowDialog();
        settingsForm=null;
    }

    // ── Inline slot editor (Apple-style) ──
    class InlineSlot {
        ComboBox combo; TextBox nameBox, txtBox;
        string[] types={"none","paste","lab","template"};
        string[] tnames={"未設定","快貼","報告","範本"};
        CardPanel card;
        public int Height { get { return card.Height; } }

        public InlineSlot(Control parent, Slot s, int idx, int top, int width) {
            bool isLab=(s.Type=="lab");
            int h=isLab?52:84;
            card=new CardPanel{Title="Ctrl+"+(idx+1),Left=12,Top=top,Width=width,Height=h};
            parent.Controls.Add(card);

            nameBox=new TextBox{Left=60,Top=36,Width=140,Text=s.Name,
                BackColor=Theme.Surface,ForeColor=Theme.Text,Font=Theme.Body,BorderStyle=BorderStyle.FixedSingle};
            card.Controls.Add(new Label{Text="名稱",Left=14,Top=39,Width=40,ForeColor=Theme.Text2,BackColor=Color.Transparent});
            card.Controls.Add(nameBox);
            combo=new ComboBox{Left=250,Top=36,Width=75,DropDownStyle=ComboBoxStyle.DropDownList,
                BackColor=Theme.Surface,ForeColor=Theme.Text,Font=Theme.Body};
            combo.Items.AddRange(tnames);combo.SelectedIndex=Array.IndexOf(types,s.Type);
            card.Controls.Add(new Label{Text="類型",Left=214,Top=39,Width=34,ForeColor=Theme.Text2,BackColor=Color.Transparent});
            card.Controls.Add(combo);

            if(!isLab){
                txtBox=new TextBox{Left=60,Top=60,Width=width-80,Height=20,
                    Text=s.Text,BackColor=Theme.Surface,ForeColor=Theme.Text,
                    Font=Theme.MonoSmall,BorderStyle=BorderStyle.FixedSingle};
                card.Controls.Add(new Label{Text="內容",Left=14,Top=63,Width=40,ForeColor=Theme.Text2,BackColor=Color.Transparent});
                card.Controls.Add(txtBox);
            } else {
                card.Controls.Add(new Label{Text="Ctrl+C 複製報告 → 此鍵貼上濃縮結果 (支援 ORD + 雲端)",
                    Left=340,Top=39,Width=280,ForeColor=Theme.Text3,BackColor=Color.Transparent,
                    Font=Theme.Small});
            }

            combo.SelectedIndexChanged+=(x,y)=>{
                int ti=combo.SelectedIndex;
                bool nowLab=(ti==2);
                if(nowLab&&txtBox!=null){txtBox.Visible=false;card.Height=52;}
                else if(!nowLab){
                    if(txtBox==null){
                        txtBox=new TextBox{Left=60,Top=60,Width=width-80,Height=20,
                            BackColor=Theme.Surface,ForeColor=Theme.Text,
                            Font=Theme.MonoSmall,BorderStyle=BorderStyle.FixedSingle};
                        card.Controls.Add(new Label{Text="內容",Left=14,Top=63,Width=40,ForeColor=Theme.Text2,BackColor=Color.Transparent});
                        card.Controls.Add(txtBox);
                    }
                    txtBox.Visible=true;card.Height=84;
                }
            };
        }
        public void SaveTo(Slot s) {
            s.Name=nameBox.Text.Trim(); if(s.Name=="")s.Name="Slot";
            s.Type=types[combo.SelectedIndex];
            if(s.Type=="paste"||s.Type=="template"){s.Text=txtBox!=null?txtBox.Text:"";s.LabItems.Clear();}
            else if(s.Type=="lab"){s.Text="";}
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
            try{File.WriteAllText(path,
                "Lab Data Formatter "+VER+"\n\n"+
                "請從設定畫面或右鍵選單開啟說明檔。\n"+
                "如果看到此訊息，請將 readme.txt 放在 exe 同資料夾。",
                Encoding.UTF8);}catch{}
        }
        try{System.Diagnostics.Process.Start(path);}catch{}
    }

    void ShowManual() {
        if(manualForm!=null&&!manualForm.IsDisposed){manualForm.BringToFront();return;}
        var f=new Form{Text="手動轉換",Size=new Size(580,460),StartPosition=FormStartPosition.CenterScreen,
            TopMost=true,BackColor=Theme.Bg,ForeColor=Theme.Text,Font=Theme.Body};
        try{int val=1;NativeMethods.DwmSetWindowAttribute(f.Handle,NativeMethods.DWMWA_USE_IMMERSIVE_DARK_MODE,ref val,sizeof(int));}catch{}
        manualForm=f;
        f.Controls.Add(new Label{Text="貼上檢驗報告或雲端藥歷：",Left=10,Top=8,Width=300,ForeColor=Theme.Text2});
        var inp=new TextBox{Left=10,Top=28,Width=545,Height=200,Multiline=true,ScrollBars=ScrollBars.Vertical,
            BackColor=Theme.Surface,ForeColor=Theme.Text,Font=Theme.MonoSmall,BorderStyle=BorderStyle.FixedSingle};
        f.Controls.Add(inp);
        var outBox=new TextBox{Left=10,Top=280,Width=545,Height=100,ReadOnly=true,Multiline=true,ScrollBars=ScrollBars.Vertical,
            BackColor=Theme.Surface,ForeColor=Theme.Accent,Font=Theme.MonoSmall,BorderStyle=BorderStyle.FixedSingle};
        f.Controls.Add(outBox);
        var b1=new PillButton{Text="轉換",Left=160,Top=238,Width=80,Height=30,IsPrimary=true};
        b1.Click+=(s,e)=>{
            var txt=inp.Text;
            string r = null;
            if (Lab.IsLabData(txt))
                r = Lab.Convert(txt, cfg.GetLabItems());
            else if (CloudLab.IsCloudLabData(txt))
                r = CloudLab.Convert(txt, cfg.GetLabItems());
            else if (CloudMed.IsCloudMedData(txt))
                r = CloudMed.Convert(txt);
            outBox.Text = r != null ? r : "未找到可辨識的資料";
            if (r != null) {
                if (CloudMed.IsCloudMedData(txt)) medResult = r;
                else labResult = r;
                SetTray(Theme.Accent,"✓ "+r.Split('\n')[0].Substring(0,Math.Min(38,r.Split('\n')[0].Length)));
                DelayYellow();
            }
        };
        var b2=new PillButton{Text="複製結果",Left=250,Top=238,Width=90,Height=30,AccentColor=Theme.Blue};
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

// ══════════ Splash Screen (Apple-style) ══════════
static class Splash {
    public static void Show() {
        var f=new Form{
            FormBorderStyle=FormBorderStyle.None,
            StartPosition=FormStartPosition.CenterScreen,
            Size=new Size(420,200),
            BackColor=Theme.Bg,
            ShowInTaskbar=false,
            TopMost=true,
            Opacity=0
        };
        Theme.SetRoundRegion(f, 16);

        var lbl1=new Label{Text="病歷小幫手",
            ForeColor=Theme.Accent,
            Font=new Font("Microsoft JhengHei UI",28,FontStyle.Bold),
            AutoSize=false,Width=420,Height=70,Top=30,
            TextAlign=ContentAlignment.MiddleCenter};
        var lbl2=new Label{Text="上線中...",
            ForeColor=Theme.Text2,
            Font=new Font("Microsoft JhengHei UI",14),
            AutoSize=false,Width=420,Height=40,Top=100,
            TextAlign=ContentAlignment.MiddleCenter};
        var lbl3=new Label{Text="Lab Data Formatter v2.0",
            ForeColor=Theme.Text3,
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
                double p=(double)tick/15; if(p>1)p=1;
                f.Opacity=p;
                if(tick>=15){phase=1;tick=0;lbl2.Text="上線完成 ✓";
                    lbl2.ForeColor=Theme.Accent;}
            } else if(phase==1){
                if(tick>=50){phase=2;tick=0;
                    cx=f.Left+f.Width/2; cy=f.Top+f.Height/2;}
            } else if(phase==2){
                double p=(double)tick/20; if(p>1)p=1;
                f.Opacity=1.0-p;
                int w=(int)(origW*(1.0-p*0.6));
                int h=(int)(origH*(1.0-p*0.6));
                if(w<10)w=10; if(h<10)h=10;
                f.SetBounds(cx-w/2,cy-h/2,w,h);
                try { Theme.SetRoundRegion(f, Math.Max((int)(8*(1.0-p)),2)); } catch {}
                if(tick>=20){timer.Stop();f.Close();}
            }
        };
        f.Shown+=delegate{timer.Start();};
        f.ShowDialog();
    }
}
