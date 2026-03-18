using System;
using System.Linq;
using System.IO;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Runtime.InteropServices;

namespace TimestampChanger
{
    // ── Custom Calendar Popup ─────────────────────────────────────────────────
    public class CalendarPopup : Form
    {
        private DateTime _selected;
        private int _year, _month;
        private bool _dark;
        public event Action<DateTime> DateSelected;

        private Color _bg, _fg, _gray, _blue, _hover, _headerBg;
        private Point _mousePos = new Point(-1, -1);

        private int _view = 0;
        private int _yearRangeStart;

        public CalendarPopup(DateTime current, bool dark)
        {
            _selected = current;
            _year = current.Year;
            _month = current.Month;
            _yearRangeStart = (_year / 12) * 12;
            _dark = dark;

            _bg       = dark ? Color.FromArgb(30,30,30)    : Color.White;
            _fg       = dark ? Color.FromArgb(230,230,230) : Color.FromArgb(30,30,30);
            _gray     = dark ? Color.FromArgb(110,110,110) : Color.FromArgb(180,180,180);
            _blue     = Color.FromArgb(0,120,212);
            _hover    = dark ? Color.FromArgb(60,60,60)    : Color.FromArgb(220,235,255);
            _headerBg = dark ? Color.FromArgb(40,40,40)    : Color.FromArgb(245,245,245);

            this.FormBorderStyle = FormBorderStyle.None;
            this.StartPosition = FormStartPosition.Manual;
            this.Size = new Size(260, 270);
            this.BackColor = _bg;
            this.TopMost = true;
            this.ShowInTaskbar = false;
            this.Deactivate += (s, e) => this.Close();
            this.DoubleBuffered = true;
        }

        private int CellW => (Width - 16) / 7;
        private int StartX => 8;
        private StringFormat SF => new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

            using (var pen = new Pen(Color.FromArgb(80,80,80), 1))
                g.DrawRectangle(pen, 0, 0, Width-1, Height-1);

            if (_view == 0) PaintDays(g);
            else if (_view == 1) PaintMonths(g);
            else PaintYears(g);
        }

        private Rectangle ArrowLeftRect  => new Rectangle(Width-56, 4, 24, 28);
        private Rectangle ArrowRightRect => new Rectangle(Width-30, 4, 24, 28);
        private Rectangle MonthRect      => new Rectangle(8,  4, 85, 28);
        private Rectangle YearRect       => new Rectangle(96, 4, 55, 28);

        private void PaintHeader(Graphics g, string title, bool showArrows = true)
        {
            g.FillRectangle(new SolidBrush(_headerBg), 0, 0, Width, 36);

            if (_view == 0)
            {
                var dt = new DateTime(_year, _month, 1);
                string monthStr = dt.ToString("MMMM");
                string yearStr  = dt.ToString("yyyy");

                bool hMonth = MonthRect.Contains(_mousePos);
                bool hYear  = YearRect.Contains(_mousePos);

                if (hMonth) g.FillRectangle(new SolidBrush(_hover), MonthRect);
                if (hYear)  g.FillRectangle(new SolidBrush(_hover), YearRect);

                g.DrawString(monthStr, new Font("Segoe UI",10f,FontStyle.Bold),
                    new SolidBrush(hMonth ? _blue : _fg), new RectangleF(MonthRect.X, MonthRect.Y, MonthRect.Width, MonthRect.Height),
                    new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center });
                g.DrawString(yearStr, new Font("Segoe UI",10f,FontStyle.Bold),
                    new SolidBrush(hYear ? _blue : _fg), new RectangleF(YearRect.X, YearRect.Y, YearRect.Width, YearRect.Height),
                    new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center });
            }
            else
            {
                bool hoverTitle = _mousePos.Y >= 4 && _mousePos.Y <= 32 && _mousePos.X >= 8 && _mousePos.X <= Width - 70;
                if (hoverTitle) g.FillRectangle(new SolidBrush(_hover), 4, 4, Width - 78, 28);
                g.DrawString(title, new Font("Segoe UI",10f,FontStyle.Bold), new SolidBrush(_fg),
                    new RectangleF(8, 4, Width-78, 28), new StringFormat { LineAlignment = StringAlignment.Center });
            }

            if (!showArrows) return;
            bool hL = ArrowLeftRect.Contains(_mousePos);
            bool hR = ArrowRightRect.Contains(_mousePos);
            if (hL) g.FillEllipse(new SolidBrush(_hover), ArrowLeftRect);
            if (hR) g.FillEllipse(new SolidBrush(_hover), ArrowRightRect);
            g.DrawString("‹", new Font("Segoe UI",13f,FontStyle.Bold), new SolidBrush(_fg), new RectangleF(ArrowLeftRect.X, ArrowLeftRect.Y, ArrowLeftRect.Width, ArrowLeftRect.Height), SF);
            g.DrawString("›", new Font("Segoe UI",13f,FontStyle.Bold), new SolidBrush(_fg), new RectangleF(ArrowRightRect.X, ArrowRightRect.Y, ArrowRightRect.Width, ArrowRightRect.Height), SF);
        }

        private void PaintDays(Graphics g)
        {
            PaintHeader(g, new DateTime(_year, _month, 1).ToString("MMMM yyyy"));

            string[] dayNames = { "Su", "Mo", "Tu", "We", "Th", "Fr", "Sa" };
            for (int i = 0; i < 7; i++)
                g.DrawString(dayNames[i], new Font("Segoe UI",8f,FontStyle.Bold), new SolidBrush(_gray),
                    new RectangleF(StartX + i * CellW, 40, CellW, 20), SF);

            var dayFont = new Font("Segoe UI", 9f);
            var boldFont = new Font("Segoe UI", 9f, FontStyle.Bold);
            var firstDay = new DateTime(_year, _month, 1);
            int startDow = (int)firstDay.DayOfWeek;
            int daysInMonth = DateTime.DaysInMonth(_year, _month);
            int prevMonth = _month == 1 ? 12 : _month - 1;
            int prevYear  = _month == 1 ? _year - 1 : _year;
            int daysInPrev = DateTime.DaysInMonth(prevYear, prevMonth);

            for (int i = 0; i < startDow; i++)
            {
                int d = daysInPrev - startDow + 1 + i;
                g.DrawString(d.ToString(), dayFont, new SolidBrush(_gray),
                    new RectangleF(StartX + i * CellW, 64, CellW, 28), SF);
            }

            int row = 0, col = startDow;
            for (int d = 1; d <= daysInMonth; d++)
            {
                var rect = new Rectangle(StartX + col * CellW, 64 + row * 32, CellW, 28);
                bool isSel     = d == _selected.Day && _month == _selected.Month && _year == _selected.Year;
                bool isToday   = d == DateTime.Today.Day && _month == DateTime.Today.Month && _year == DateTime.Today.Year;
                bool isHovered = col == (_mousePos.X - StartX) / CellW && row == (_mousePos.Y - 64) / 32 && _mousePos.Y >= 64;

                if (isSel)
                {
                    g.FillEllipse(new SolidBrush(_blue), rect.X+2, rect.Y+1, rect.Width-4, rect.Height-2);
                    g.DrawString(d.ToString(), boldFont, Brushes.White, new RectangleF(rect.X,rect.Y,rect.Width,rect.Height), SF);
                }
                else if (isToday)
                {
                    if (isHovered) g.FillEllipse(new SolidBrush(_hover), rect.X+2, rect.Y+1, rect.Width-4, rect.Height-2);
                    using (var pen = new Pen(_blue, 1.5f))
                        g.DrawEllipse(pen, rect.X+2, rect.Y+1, rect.Width-4, rect.Height-2);
                    g.DrawString(d.ToString(), boldFont, new SolidBrush(_blue), new RectangleF(rect.X,rect.Y,rect.Width,rect.Height), SF);
                }
                else
                {
                    if (isHovered) g.FillEllipse(new SolidBrush(_hover), rect.X+2, rect.Y+1, rect.Width-4, rect.Height-2);
                    g.DrawString(d.ToString(), dayFont, new SolidBrush(_fg), new RectangleF(rect.X,rect.Y,rect.Width,rect.Height), SF);
                }
                col++; if (col == 7) { col = 0; row++; }
            }

            int nextDay = 1;
            while (col != 0)
            {
                g.DrawString(nextDay.ToString(), dayFont, new SolidBrush(_gray),
                    new RectangleF(StartX + col * CellW, 64 + row * 32, CellW, 28), SF);
                nextDay++; col++; if (col == 7) { col = 0; row++; }
            }
        }

        private void PaintMonths(Graphics g)
        {
            PaintHeader(g, _year.ToString(), showArrows: true);
            string[] months = { "Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec" };
            int cellW = (Width - 16) / 3;
            int cellH = (Height - 44) / 4;
            for (int i = 0; i < 12; i++)
            {
                int rx = StartX + (i % 3) * cellW;
                int ry = 44 + (i / 3) * cellH;
                var rect = new Rectangle(rx, ry, cellW, cellH);
                bool isSel = i + 1 == _month;
                bool isHov = rect.Contains(_mousePos);
                if (isSel) g.FillRectangle(new SolidBrush(_blue), rect.X+2, rect.Y+2, rect.Width-4, rect.Height-4);
                else if (isHov) g.FillRectangle(new SolidBrush(_hover), rect.X+2, rect.Y+2, rect.Width-4, rect.Height-4);
                g.DrawString(months[i], new Font("Segoe UI", 9f, isSel ? FontStyle.Bold : FontStyle.Regular),
                    new SolidBrush(isSel ? Color.White : _fg), new RectangleF(rect.X,rect.Y,rect.Width,rect.Height), SF);
            }
        }

        private void PaintYears(Graphics g)
        {
            PaintHeader(g, $"{_yearRangeStart} – {_yearRangeStart+11}", showArrows: true);
            int cellW = (Width - 16) / 3;
            int cellH = (Height - 44) / 4;
            for (int i = 0; i < 12; i++)
            {
                int yr = _yearRangeStart + i;
                int rx = StartX + (i % 3) * cellW;
                int ry = 44 + (i / 3) * cellH;
                var rect = new Rectangle(rx, ry, cellW, cellH);
                bool isSel = yr == _year;
                bool isHov = rect.Contains(_mousePos);
                if (isSel) g.FillRectangle(new SolidBrush(_blue), rect.X+2, rect.Y+2, rect.Width-4, rect.Height-4);
                else if (isHov) g.FillRectangle(new SolidBrush(_hover), rect.X+2, rect.Y+2, rect.Width-4, rect.Height-4);
                g.DrawString(yr.ToString(), new Font("Segoe UI", 9f, isSel ? FontStyle.Bold : FontStyle.Regular),
                    new SolidBrush(isSel ? Color.White : _fg), new RectangleF(rect.X,rect.Y,rect.Width,rect.Height), SF);
            }
        }

        protected override void OnMouseClick(MouseEventArgs e)
        {
            if (e.Y < 36)
            {
                if (_view == 0 && MonthRect.Contains(e.Location)) { _view = 1; Refresh(); return; }
                if (_view == 0 && YearRect.Contains(e.Location))  { _view = 2; _yearRangeStart = (_year/12)*12; Refresh(); return; }
                if (_view == 1 && e.X >= 8 && e.X <= Width-70)    { _view = 2; _yearRangeStart = (_year/12)*12; Refresh(); return; }

                bool left  = ArrowLeftRect.Contains(e.Location);
                bool right = ArrowRightRect.Contains(e.Location);
                if (_view == 0)
                {
                    if (left)  { if (_month==1){_month=12;_year--;} else _month--; }
                    if (right) { if (_month==12){_month=1;_year++;} else _month++; }
                }
                else if (_view == 1)
                {
                    if (left)  _year--;
                    if (right) _year++;
                }
                else
                {
                    if (left)  _yearRangeStart -= 12;
                    if (right) _yearRangeStart += 12;
                }
                Refresh(); return;
            }

            if (_view == 1)
            {
                int cellW = (Width - 16) / 3;
                int cellH = (Height - 44) / 4;
                int col = (e.X - StartX) / cellW;
                int row = (e.Y - 44) / cellH;
                int idx = row * 3 + col;
                if (idx >= 0 && idx < 12) { _month = idx + 1; _view = 0; Refresh(); }
                return;
            }

            if (_view == 2)
            {
                int cellW = (Width - 16) / 3;
                int cellH = (Height - 44) / 4;
                int col = (e.X - StartX) / cellW;
                int row = (e.Y - 44) / cellH;
                int idx = row * 3 + col;
                if (idx >= 0 && idx < 12) { _year = _yearRangeStart + idx; _view = 0; Refresh(); }
                return;
            }

            if (_view == 0 && e.Y >= 64)
            {
                int col = (e.X - StartX) / CellW;
                int row = (e.Y - 64) / 32;
                int startDow = (int)new DateTime(_year, _month, 1).DayOfWeek;
                int day = row * 7 + col - startDow + 1;
                int daysInMonth = DateTime.DaysInMonth(_year, _month);
                if (day >= 1 && day <= daysInMonth)
                {
                    _selected = new DateTime(_year, _month, day);
                    DateSelected?.Invoke(_selected);
                    this.Close();
                }
            }
        }

        protected override void OnMouseMove(MouseEventArgs e) { _mousePos = e.Location; Refresh(); }
        protected override void OnMouseLeave(EventArgs e) { _mousePos = new Point(-1,-1); Refresh(); }

        protected override CreateParams CreateParams
        {
            get { var cp = base.CreateParams; cp.ClassStyle |= 0x20000; return cp; }
        }
    }

    // ── Smart Spinner ─────────────────────────────────────────────────────────
    public class SmartSpinner : Control
    {
        private int _value, _min, _max;
        private bool _dark;
        private string _typing = "";
        private bool _cleared = false;
        private bool _caretVisible = false;
        private System.Windows.Forms.Timer _caretTimer;
        public event EventHandler ValueChanged;

        public SmartSpinner(bool dark, int min, int max, int val = 0)
        {
            _dark = dark; _min = min; _max = max; _value = val;
            this.SetStyle(ControlStyles.Selectable | ControlStyles.UserPaint |
                          ControlStyles.AllPaintingInWmPaint | ControlStyles.DoubleBuffer, true);
            this.Cursor = Cursors.IBeam;
            this.TabStop = true;
            _caretTimer = new System.Windows.Forms.Timer { Interval = 530 };
            _caretTimer.Tick += (s,e) => { _caretVisible = !_caretVisible; Invalidate(); };
        }

        public int Value
        {
            get => _value;
            set { _value = Clamp(value); Invalidate(); ValueChanged?.Invoke(this, EventArgs.Empty); }
        }

        public void SetRange(int min, int max) { _min = min; _max = max; _value = Clamp(_value); Invalidate(); }

        private int Clamp(int v) => Math.Max(_min, Math.Min(_max, v));

        private Color Bg     => _dark ? Color.FromArgb(50,50,50)    : Color.White;
        private Color Fg     => _dark ? Color.FromArgb(220,220,220) : Color.Black;
        private Color BtnBg  => _dark ? Color.FromArgb(65,65,65)    : Color.FromArgb(225,225,225);
        private Color BtnHov => _dark ? Color.FromArgb(85,85,85)    : Color.FromArgb(200,215,235);
        private Color Border => _dark ? Color.FromArgb(80,80,80)    : Color.FromArgb(180,180,180);

        private Rectangle UpRect   => new Rectangle(Width-16, 0,        16, Height/2);
        private Rectangle DownRect => new Rectangle(Width-16, Height/2, 16, Height - Height/2);
        private Rectangle TextRect => new Rectangle(0, 0, Width-16, Height);

        private bool _hoverUp, _hoverDown;

        private string DisplayText => _cleared ? "" : (_typing.Length > 0 ? _typing : _value.ToString());

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            g.Clear(Bg);

            using (var pen = new Pen(this.Focused ? Color.FromArgb(0,120,212) : Border))
                g.DrawRectangle(pen, 0, 0, Width-1, Height-1);

            var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
            g.DrawString(DisplayText, new Font("Segoe UI", 9f), new SolidBrush(Fg),
                new RectangleF(0, 0, Width-16, Height), sf);

            if (this.Focused && _caretVisible)
            {
                var size = g.MeasureString(DisplayText, new Font("Segoe UI", 9f));
                int cx = (Width-16)/2 + (int)(size.Width/2);
                using (var pen = new Pen(Fg, 1.5f))
                    g.DrawLine(pen, cx, 3, cx, Height-4);
            }

            using (var pen = new Pen(Border))
                g.DrawLine(pen, Width-16, 0, Width-16, Height);

            g.FillRectangle(new SolidBrush(_hoverUp   ? BtnHov : BtnBg), UpRect);
            g.FillRectangle(new SolidBrush(_hoverDown ? BtnHov : BtnBg), DownRect);
            using (var pen = new Pen(Border))
                g.DrawLine(pen, Width-16, Height/2, Width-1, Height/2);

            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            DrawArrow(g, UpRect, true);
            DrawArrow(g, DownRect, false);
        }

        private void DrawArrow(Graphics g, Rectangle r, bool up)
        {
            int cx = r.X + r.Width/2, cy = r.Y + r.Height/2;
            var pts = up
                ? new[]{ new Point(cx-3,cy+1), new Point(cx,cy-2), new Point(cx+3,cy+1) }
                : new[]{ new Point(cx-3,cy-1), new Point(cx,cy+2), new Point(cx+3,cy-1) };
            using (var pen = new Pen(_dark ? Color.FromArgb(220,220,220) : Color.Black, 1.5f))
                g.DrawLines(pen, pts);
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            this.Focus();
            if (UpRect.Contains(e.Location))   { CommitFromOutside(); Value++; return; }
            if (DownRect.Contains(e.Location)) { CommitFromOutside(); Value--; return; }
            _typing = ""; _cleared = false;
            _caretVisible = true; _caretTimer.Start();
            Invalidate();
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            bool hu = UpRect.Contains(e.Location), hd = DownRect.Contains(e.Location);
            if (hu != _hoverUp || hd != _hoverDown) { _hoverUp = hu; _hoverDown = hd; Invalidate(); }
            this.Cursor = TextRect.Contains(e.Location) ? Cursors.IBeam : Cursors.Default;
        }

        protected override void OnMouseLeave(EventArgs e) { _hoverUp = _hoverDown = false; Invalidate(); }

        protected override void OnKeyDown(KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Up)   { CommitFromOutside(); Value++; e.Handled=true; return; }
            if (e.KeyCode == Keys.Down) { CommitFromOutside(); Value--; e.Handled=true; return; }
            if (e.KeyCode == Keys.Back)
            {
                if (_cleared) { e.Handled=true; return; }
                if (_typing.Length == 0) _typing = _value.ToString();
                _typing = _typing.Substring(0, _typing.Length-1);
                if (_typing.Length == 0) { _cleared = true; }
                else { _value = Clamp(int.Parse(_typing)); ValueChanged?.Invoke(this, EventArgs.Empty); }
                Invalidate(); e.Handled=true; return;
            }
        }

        protected override void OnKeyPress(KeyPressEventArgs e)
        {
            if (!char.IsDigit(e.KeyChar)) { e.Handled=true; return; }
            int maxLen = _max >= 100 ? 3 : 2;
            if (_typing.Length >= maxLen) { e.Handled=true; return; }
            _cleared = false;
            _typing += e.KeyChar;
            _value = Clamp(int.Parse(_typing));
            ValueChanged?.Invoke(this, EventArgs.Empty);
            if (_typing.Length == maxLen) { CommitTyping(); }
            Invalidate(); e.Handled=true;
        }

        private void CommitTyping()
        {
            if (_typing.Length > 0) { _value = Clamp(int.Parse(_typing)); _typing = ""; }
        }

        public void CommitFromOutside()
        {
            if (_cleared) { _value = _min; _cleared = false; }
            CommitTyping();
            _caretTimer.Stop(); _caretVisible = false;
            _typing = "";
            this.Parent?.Focus();
            Invalidate();
        }

        protected override void OnGotFocus(EventArgs e)  { _caretVisible=true; _caretTimer.Start(); Invalidate(); }
        protected override void OnLostFocus(EventArgs e) { if (_cleared) { _value=_min; _cleared=false; } CommitTyping(); _caretTimer.Stop(); _caretVisible=false; Invalidate(); }
        protected override bool IsInputKey(Keys k) => k==Keys.Up || k==Keys.Down || base.IsInputKey(k);
    }

    // ── Smart Date Input ──────────────────────────────────────────────────────
    // Uses a real TextBox for full native drag-selection, styled to match the app.
    public class SmartDateBox : UserControl
    {
        private TextBox _tb;
        private bool _dark;
        private bool _updating = false;
        public event EventHandler DateChanged;

        public SmartDateBox(bool dark)
        {
            _dark = dark;
            this.SetStyle(ControlStyles.Selectable, true);
            this.TabStop = false;

            _tb = new TextBox
            {
                BorderStyle = BorderStyle.None,
                Font        = new Font("Segoe UI", 11f),
                BackColor   = dark ? Color.FromArgb(50,50,50)    : Color.White,
                ForeColor   = dark ? Color.FromArgb(220,220,220) : Color.Black,
                TextAlign   = HorizontalAlignment.Left,
                Dock        = DockStyle.Fill,
                TabStop     = true,
            };

            this.BackColor = dark ? Color.FromArgb(50,50,50) : Color.White;
            this.Padding   = new Padding(3, 2, 0, 0);
            this.Controls.Add(_tb);

            _tb.TextChanged += (s,e) => { if (!_updating) DateChanged?.Invoke(this, EventArgs.Empty); };

            // Track which segment the mouse went down in, clamp selection to that segment
            int _mouseDownSeg = -1;
            _tb.MouseDown += (s, e) =>
            {
                _mouseDownSeg = GetSegAtX(e.X);
            };
            _tb.MouseUp += (s, e) => ClampSelection(_mouseDownSeg);
            _tb.MouseMove += (s, e) =>
            {
                if (e.Button == MouseButtons.Left) ClampSelection(_mouseDownSeg);
            };
            _tb.MouseDoubleClick += (s, e) =>
            {
                int seg = GetSegAtX(e.X);
                var parts = _tb.Text.Split('/');
                int mLen = parts.Length > 0 ? parts[0].Length : 0;
                int dLen = parts.Length > 1 ? parts[1].Length : 0;
                if      (seg == 0) { _tb.SelectionStart = 0;                   _tb.SelectionLength = mLen; }
                else if (seg == 1) { _tb.SelectionStart = mLen + 1;            _tb.SelectionLength = dLen; }
                else               { _tb.SelectionStart = mLen + 1 + dLen + 1; _tb.SelectionLength = _tb.Text.Length - (mLen + 1 + dLen + 1); }
            };

            // Only allow digits and slashes
            _tb.KeyPress += (s, e) =>
            {
                if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && e.KeyChar != '/')
                { e.Handled = true; return; }

                if (char.IsControl(e.KeyChar)) return;

                // Build what the text would look like after this keypress
                int caret = _tb.SelectionStart;
                string current = _tb.Text;
                string next = current.Substring(0, caret) + e.KeyChar + current.Substring(caret + _tb.SelectionLength);

                // Parse the parts: month / day / year
                var parts = next.Split('/');
                bool editingMonth = caret <= (parts.Length > 0 ? parts[0].Length : 0);

                if (parts.Length >= 1 && parts[0].Length > 0)
                {
                    if (parts[0] == "0") { e.Handled = true; return; }
                    if (int.TryParse(parts[0], out int m) && m > 12) { e.Handled = true; return; }
                    if (parts[0].Length > 2) { e.Handled = true; return; }
                }
                if (parts.Length >= 2 && parts[1].Length > 0)
                {
                    if (parts[1] == "0") { e.Handled = true; return; }
                    if (!editingMonth) // only check day limit when actually editing the day
                    {
                        int parsedMonth = (parts.Length >= 1 && int.TryParse(parts[0], out int pm)) ? Math.Max(1,Math.Min(12,pm)) : DateTime.Today.Month;
                        int parsedYear  = (parts.Length >= 3 && int.TryParse(parts[2], out int py)) ? Math.Max(1900,Math.Min(DateTime.Today.Year,py)) : DateTime.Today.Year;
                        if (int.TryParse(parts[1], out int d) && d > DateTime.DaysInMonth(parsedYear, parsedMonth)) { e.Handled = true; return; }
                    }
                    if (parts[1].Length > 2) { e.Handled = true; return; }
                }
                if (parts.Length >= 3 && parts[2].Length > 0)
                {
                    if (int.TryParse(parts[2], out int y) && y > DateTime.Today.Year) { e.Handled = true; return; }
                    if (parts[2].Length > 4) { e.Handled = true; return; }
                }
            };

            // Draw border ourselves so we can color it blue on focus
            this.Paint += (s, ev) =>
            {
                var borderColor = _tb.Focused
                    ? Color.FromArgb(0,120,212)
                    : (dark ? Color.FromArgb(80,80,80) : Color.FromArgb(180,180,180));
                using (var pen = new Pen(borderColor, 1))
                    ev.Graphics.DrawRectangle(pen, 0, 0, Width-1, Height-1);
            };
            _tb.GotFocus  += (s,e) => Invalidate();
            _tb.LostFocus += (s,e) => { Invalidate(); ParseAndFix(); };
        }

        public new bool Enabled
        {
            get => _tb.Enabled;
            set { _tb.Enabled = value; base.Enabled = value; }
        }

        public DateTime Value
        {
            get
            {
                if (TryParse(_tb.Text, out var dt)) return dt;
                return DateTime.Today;
            }
            set
            {
                _updating = true;
                _tb.Text = $"{value.Month}/{value.Day}/{value.Year}";
                _updating = false;
            }
        }

        private bool TryParse(string text, out DateTime result)
        {
            return DateTime.TryParseExact(text,
                new[]{ "M/d/yyyy","M/d/yy","M/dd/yyyy","MM/dd/yyyy","M/d/yyy" },
                System.Globalization.CultureInfo.InvariantCulture,
                System.Globalization.DateTimeStyles.None, out result);
        }

        private void ParseAndFix()
        {
            // Parse parts directly from raw text so invalid dates like 2/31 don't fall back to today
            var parts = _tb.Text.Split('/');
            int year  = DateTime.Today.Year;
            int month = DateTime.Today.Month;
            int day   = DateTime.Today.Day;

            if (parts.Length >= 1 && int.TryParse(parts[0], out int m)) month = m;
            if (parts.Length >= 2 && int.TryParse(parts[1], out int d)) day   = d;
            if (parts.Length >= 3 && int.TryParse(parts[2], out int y)) year  = y;

            // Clamp everything
            year  = Math.Max(1900, Math.Min(DateTime.Today.Year, year));
            month = Math.Max(1, Math.Min(12, month));
            day   = Math.Max(1, Math.Min(DateTime.DaysInMonth(year, month), day)); // clamps 31→28 for Feb etc

            _updating = true;
            _tb.Text = $"{month}/{day}/{year}";
            _updating = false;
        }

        // Returns segment based on pixel X position
        private int GetSegAtX(int x)
        {
            // Use GetCharIndexFromPosition to find char at pixel, then use GetSegAt
            int charIdx = _tb.GetCharIndexFromPosition(new Point(x, Height / 2));
            return GetSegAt(charIdx);
        }

        // Returns which segment (0=month,1=day,2=year) a character index belongs to, or -1 if on a slash
        private int GetSegAt(int charIdx)
        {
            var parts = _tb.Text.Split('/');
            int m = parts.Length > 0 ? parts[0].Length : 0;
            int d = parts.Length > 1 ? parts[1].Length : 0;
            if (charIdx <= m) return 0;
            if (charIdx == m + 1) return -1; // on first slash
            if (charIdx <= m + 1 + d) return 1;
            if (charIdx == m + 1 + d + 1) return -1; // on second slash
            return 2;
        }

        // Clamps the textbox selection so it can't cross slash separators
        private void ClampSelection(int anchorSeg)
        {
            if (anchorSeg < 0) return;
            var parts = _tb.Text.Split('/');
            int mLen = parts.Length > 0 ? parts[0].Length : 0;
            int dLen = parts.Length > 1 ? parts[1].Length : 0;

            int segStart, segEnd;
            if (anchorSeg == 0)      { segStart = 0;              segEnd = mLen; }
            else if (anchorSeg == 1) { segStart = mLen + 1;       segEnd = mLen + 1 + dLen; }
            else                     { segStart = mLen + 1 + dLen + 1; segEnd = _tb.Text.Length; }

            int selStart = _tb.SelectionStart;
            int selEnd   = selStart + _tb.SelectionLength;

            int clampedStart = Math.Max(segStart, Math.Min(segEnd, selStart));
            int clampedEnd   = Math.Max(segStart, Math.Min(segEnd, selEnd));

            if (clampedStart != selStart || (clampedEnd - clampedStart) != _tb.SelectionLength)
            {
                _tb.SelectionStart  = clampedStart;
                _tb.SelectionLength = clampedEnd - clampedStart;
            }
        }

        public void CommitFromOutside() { ParseAndFix(); }
    }

    // ── Timestamp Row ─────────────────────────────────────────────────────────
    public class TimestampRow : UserControl
    {
        private CheckBox chk;
        private Button dateBtn;
        private SmartDateBox _dateBox;
        private SmartSpinner nudH, nudM, nudS;
        private DateTime _currentDate = DateTime.Today;
        private bool _dark;
        private bool _use24h = true;

        public bool IsEnabled => chk.Checked;
        public Label LabelControl { get; private set; }

        public TimestampRow(string label, bool dark)
        {
            _dark = dark;
            var bg = dark ? Color.FromArgb(32,32,32) : Color.FromArgb(240,240,240);
            var fg = dark ? Color.FromArgb(220,220,220) : Color.Black;

            this.Width = 450; this.Height = 30;
            this.BackColor = bg;

            chk = new CheckBox { Checked=true, Left=0, Top=5, Width=20, Height=20, BackColor=bg, ForeColor=fg };
            chk.CheckedChanged += (s,e) => UpdateState();

            LabelControl = new Label { Text=label, Left=26, Top=4, Width=110, Height=22,
                ForeColor=fg, BackColor=bg, TextAlign=ContentAlignment.MiddleLeft };

            dateBtn = new Button
            {
                Left=240, Top=2, Width=26, Height=26,
                FlatStyle=FlatStyle.Flat,
                BackColor=dark ? Color.FromArgb(50,50,50) : Color.FromArgb(230,230,230),
                ForeColor=fg, Text="📅", Cursor=Cursors.Hand,
                Font=new Font("Segoe UI", 9f),
            };
            dateBtn.FlatAppearance.BorderColor = dark ? Color.FromArgb(80,80,80) : Color.FromArgb(180,180,180);
            dateBtn.FlatAppearance.BorderSize = 1;
            dateBtn.Click += OpenCalendar;

            _dateBox = new SmartDateBox(dark) { Left=140, Top=2, Width=100, Height=26 };
            _dateBox.DateChanged += (s,e) => _currentDate = _dateBox.Value;

            nudH = MakeNud(281, 23);
            nudM = MakeNud(327, 59);
            nudS = MakeNud(373, 59);

            var c1 = new Label { Text=":", Left=321, Top=5, Width=8, Height=20, ForeColor=fg, BackColor=bg, TextAlign=ContentAlignment.MiddleCenter };
            var c2 = new Label { Text=":", Left=367, Top=5, Width=8, Height=20, ForeColor=fg, BackColor=bg, TextAlign=ContentAlignment.MiddleCenter };

            this.Controls.AddRange(new Control[]{ chk, LabelControl, _dateBox, dateBtn, nudH, c1, nudM, c2, nudS });
        }

        private void OpenCalendar(object? sender, EventArgs e)
        {
            if (!chk.Checked) return;
            var cal = new CalendarPopup(_currentDate, _dark);
            cal.DateSelected += date => { _currentDate = date; _dateBox.Value = date; };
            var screenPos = dateBtn.PointToScreen(new Point(0, dateBtn.Height));
            cal.Location = screenPos;
            cal.Show();
        }

        private SmartSpinner MakeNud(int x, int max) =>
            new SmartSpinner(_dark, 0, max) { Left=x, Top=2, Width=40, Height=26 };

        // AM/PM label shown in 12h mode
        private Label? _amPmLabel;

        public void SetTimeFormat(bool use24h)
        {
            _use24h = use24h;
            int currentHour24 = _use24h
                ? nudH.Value
                : (nudH.Value % 12) + (_amPmLabel?.Text == "PM" ? 12 : 0);

            if (use24h)
            {
                nudH.SetRange(0, 23);
                nudH.Value = currentHour24;
                if (_amPmLabel != null) { this.Controls.Remove(_amPmLabel); _amPmLabel = null; }
            }
            else
            {
                nudH.SetRange(1, 12);
                bool isPm = currentHour24 >= 12;
                int h12 = currentHour24 % 12; if (h12 == 0) h12 = 12;
                nudH.Value = h12;
                if (_amPmLabel == null)
                {
                    var fg = _dark ? Color.FromArgb(220,220,220) : Color.Black;
                    var bg = _dark ? Color.FromArgb(32,32,32) : Color.FromArgb(240,240,240);
                    _amPmLabel = new Label
                    {
                        Left=410, Top=5, Width=30, Height=20,
                        Text= isPm ? "PM" : "AM",
                        ForeColor=fg, BackColor=bg,
                        TextAlign=ContentAlignment.MiddleCenter,
                        Font=new Font("Segoe UI", 7.5f, FontStyle.Bold),
                        Cursor=Cursors.Hand,
                    };
                    _amPmLabel.Click += (s,e) => { _amPmLabel.Text = _amPmLabel.Text == "AM" ? "PM" : "AM"; };
                    this.Controls.Add(_amPmLabel);
                }
                else _amPmLabel.Text = isPm ? "PM" : "AM";
            }
        }

        public void SetDateTime(DateTime dt)
        {
            _currentDate = dt;
            _dateBox.Value = dt;
            if (_use24h)
            {
                nudH.Value = dt.Hour;
            }
            else
            {
                bool isPm = dt.Hour >= 12;
                int h12 = dt.Hour % 12; if (h12 == 0) h12 = 12;
                nudH.Value = h12;
                if (_amPmLabel != null) _amPmLabel.Text = isPm ? "PM" : "AM";
            }
            nudM.Value = dt.Minute; nudS.Value = dt.Second;
        }

        public DateTime GetDateTime()
        {
            _currentDate = _dateBox.Value;
            int hour = nudH.Value;
            if (!_use24h)
            {
                bool isPm = _amPmLabel?.Text == "PM";
                if (isPm && hour < 12) hour += 12;
                if (!isPm && hour == 12) hour = 0;
            }
            return new DateTime(_currentDate.Year, _currentDate.Month, _currentDate.Day,
                hour, nudM.Value, nudS.Value);
        }

        public void SetEnabled(bool enabled) { chk.Checked = enabled; UpdateState(); }

        private void UpdateState()
        {
            dateBtn.Enabled = chk.Checked;
            _dateBox.Enabled = chk.Checked;
            nudH.Enabled = nudM.Enabled = nudS.Enabled = chk.Checked;
        }
    }

    // ── Main Form ─────────────────────────────────────────────────────────────
    public partial class Form1 : Form, IMessageFilter
    {
        private string[] _filePaths;
        private TimestampRow rowCreated, rowModified, rowAccessed;
        private Button btnTouch, btnCancel, btnOK, btnApply;
        private ToolTip toolTip;
        private bool _dark;
        private bool _use24h = true;
        private Button _btnTimeFormat;

        // Dark title bar
        [DllImport("dwmapi.dll")]
        private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);
        private const int DWMWA_USE_IMMERSIVE_DARK_MODE = 20;

        public Form1(string[] filePaths)
        {
            _filePaths = filePaths;
            BuildUI();
            LoadTimestamps();
        }

        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);
            if (_dark)
            {
                int val = 1;
                DwmSetWindowAttribute(this.Handle, DWMWA_USE_IMMERSIVE_DARK_MODE, ref val, sizeof(int));
            }
        }

        private void BuildUI()
        {
            Application.AddMessageFilter(this);
            _dark = IsDarkMode();
            var bg   = _dark ? Color.FromArgb(32,32,32)    : Color.FromArgb(240,240,240);
            var fg   = _dark ? Color.FromArgb(220,220,220) : Color.Black;
            var card = _dark ? Color.FromArgb(45,45,45)    : Color.White;
            var blue = Color.FromArgb(0,120,212);

            this.Text = "Timestamp Changer";
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false; this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.ClientSize = new Size(450, 290);
            this.BackColor = bg;
            this.Font = new Font("Segoe UI", 9f);
            toolTip = new ToolTip();

            // File panel
            var filePanel = new Panel { Left=12, Top=12, Width=426, Height=34, BackColor=card };
            string fileLabel = _filePaths.Length == 1
                ? "📄  " + Path.GetFileName(_filePaths[0])
                : "📄  " + _filePaths.Length + " files selected";
            var lblFile = new Label { Text=fileLabel,
                Font=new Font("Segoe UI",9f,FontStyle.Bold), ForeColor=fg, BackColor=card,
                Left=8, Top=0, Width=410, Height=34, TextAlign=ContentAlignment.MiddleLeft };
            filePanel.Controls.Add(lblFile);
            this.Controls.Add(filePanel);

            // Column headers
            int hY = 56;
            var grayFont = new Font("Segoe UI", 8f);
            AddLabel(this, "Date",  Color.Gray, grayFont, 155, hY, 80,  18);
            AddLabel(this, "HH",    Color.Gray, grayFont, 285, hY, 40,  18, ContentAlignment.MiddleCenter);
            AddLabel(this, "MM",    Color.Gray, grayFont, 331, hY, 40,  18, ContentAlignment.MiddleCenter);
            AddLabel(this, "SS",    Color.Gray, grayFont, 377, hY, 40,  18, ContentAlignment.MiddleCenter);

            // 24h/12h toggle button
            _btnTimeFormat = new Button
            {
                Text = "24h", Left = 419, Top = hY - 2, Width = 28, Height = 18,
                FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 6.5f, FontStyle.Bold),
                BackColor = blue, ForeColor = Color.White, Cursor = Cursors.Hand,
            };
            _btnTimeFormat.FlatAppearance.BorderSize = 0;
            _btnTimeFormat.Click += (s, e) =>
            {
                _use24h = !_use24h;
                _btnTimeFormat.Text = _use24h ? "24h" : "12h";
                _btnTimeFormat.BackColor = _use24h ? blue : Color.FromArgb(160, 0, 80);
                rowCreated.SetTimeFormat(_use24h);
                rowModified.SetTimeFormat(_use24h);
                rowAccessed.SetTimeFormat(_use24h);
            };
            this.Controls.Add(_btnTimeFormat);

            // Rows
            rowCreated  = new TimestampRow("Created",       _dark); rowCreated.Location  = new Point(12, 76);
            rowModified = new TimestampRow("Last Modified",  _dark); rowModified.Location = new Point(12, 116);
            rowAccessed = new TimestampRow("Last Accessed",  _dark); rowAccessed.Location = new Point(12, 156);
            this.Controls.AddRange(new Control[]{ rowCreated, rowModified, rowAccessed });

            toolTip.SetToolTip(rowCreated.LabelControl,  "When the file was first created.");
            toolTip.SetToolTip(rowModified.LabelControl, "When the file content was last changed. This is what most apps and Explorer show.");
            toolTip.SetToolTip(rowAccessed.LabelControl, "When the file was last opened or read.");

            // Divider
            var div = new Panel { Left=12, Top=196, Width=426, Height=1, BackColor=Color.FromArgb(100,100,100) };
            this.Controls.Add(div);

            // Buttons
            btnTouch  = MakeBtn("⚡ Touch", 90, 30, _dark ? Color.FromArgb(55,55,55) : Color.FromArgb(225,225,225), fg);
            btnCancel = MakeBtn("Cancel",   75, 30, _dark ? Color.FromArgb(55,55,55) : Color.FromArgb(225,225,225), fg);
            btnOK     = MakeBtn("OK",       75, 30, _dark ? Color.FromArgb(55,55,55) : Color.FromArgb(225,225,225), fg);
            btnApply  = MakeBtn("Apply",    75, 30, blue, Color.White);

            btnTouch.Location  = new Point(12,  248);
            btnCancel.Location = new Point(192, 248);
            btnOK.Location     = new Point(275, 248);
            btnApply.Location  = new Point(358, 248);

            toolTip.SetToolTip(btnTouch, "Set Last Modified to the current date & time");
            btnTouch.Click  += (s,e) => Touch();
            btnCancel.Click += (s,e) => Close();
            btnOK.Click     += (s,e) => { Apply(silent: true); Close(); };
            btnApply.Click  += (s,e) => Apply(silent: false);

            this.Controls.AddRange(new Control[]{ btnTouch, btnCancel, btnOK, btnApply });
        }

        private Label AddLabel(Control parent, string text, Color fg, Font font, int x, int y, int w, int h,
            ContentAlignment align = ContentAlignment.MiddleLeft)
        {
            var lbl = new Label { Text=text, ForeColor=fg, Font=font,
                Left=x, Top=y, Width=w, Height=h, TextAlign=align, BackColor=Color.Transparent };
            parent.Controls.Add(lbl);
            return lbl;
        }

        private Button MakeBtn(string text, int w, int h, Color bg, Color fg)
        {
            var btn = new Button { Text=text, Width=w, Height=h, FlatStyle=FlatStyle.Flat,
                BackColor=bg, ForeColor=fg, Cursor=Cursors.Hand };
            btn.FlatAppearance.BorderSize = 0;
            return btn;
        }

        private bool IsDarkMode()
        {
            try
            {
                var key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize");
                var val = key?.GetValue("AppsUseLightTheme");
                return val is int i && i == 0;
            }
            catch { return false; }
        }

        private void LoadTimestamps()
        {
            try
            {
                rowCreated.SetDateTime(File.GetCreationTime(_filePaths[0]));
                rowModified.SetDateTime(File.GetLastWriteTime(_filePaths[0]));
                rowAccessed.SetDateTime(File.GetLastAccessTime(_filePaths[0]));
            }
            catch (Exception ex)
            {
                MessageBox.Show("Could not read timestamps:\n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Close();
            }
        }

        public bool PreFilterMessage(ref Message m)
        {
            const int WM_LBUTTONDOWN = 0x0201;
            if (m.Msg == WM_LBUTTONDOWN)
            {
                var ctrl = Control.FromHandle(m.HWnd);
                if (ctrl != null && !(ctrl is SmartDateBox) && !(ctrl is SmartSpinner) && !(ctrl?.Parent is SmartDateBox))
                {
                    var form = ctrl.FindForm();
                    if (form is CalendarPopup) return false;
                    foreach (Control row in new Control[]{ rowCreated, rowModified, rowAccessed })
                        foreach (Control c in row.Controls)
                        {
                            if (c is SmartDateBox sdb) sdb.CommitFromOutside();
                            if (c is SmartSpinner ss) ss.CommitFromOutside();
                        }
                }
            }
            return false;
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            Application.RemoveMessageFilter(this);
            base.OnFormClosed(e);
        }

        private void Touch() { rowModified.SetDateTime(DateTime.Now); rowModified.SetEnabled(true); }

        private void Apply(bool silent = false)
        {
            int failed = 0;
            foreach (var path in _filePaths)
            {
                try
                {
                    if (rowCreated.IsEnabled)  File.SetCreationTime(path,   rowCreated.GetDateTime());
                    if (rowModified.IsEnabled) File.SetLastWriteTime(path,  rowModified.GetDateTime());
                    if (rowAccessed.IsEnabled) File.SetLastAccessTime(path, rowAccessed.GetDateTime());
                }
                catch { failed++; }
            }
            if (!silent && failed == 0)
                MessageBox.Show(_filePaths.Length == 1 ? "Timestamps updated successfully!" : $"Updated {_filePaths.Length} files successfully!", "Done", MessageBoxButtons.OK, MessageBoxIcon.Information);
            else if (failed > 0)
                MessageBox.Show($"{failed} file(s) could not be updated.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }

    // ── Entry point ───────────────────────────────────────────────────────────
    static class Program
    {
        private const string PipeName  = "TimestampChangerPipe";
        private const string MutexName = "TimestampChangerMutex";

        [STAThread]
        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            if (args.Length < 1) { MessageBox.Show("No file specified.\nRight-click a file and choose 'Change Timestamps'."); return; }
            var path = args.Where(a => File.Exists(a) || Directory.Exists(a)).FirstOrDefault();
            if (path == null) { MessageBox.Show("No valid files found."); return; }

            bool createdNew;
            var mutex = new System.Threading.Mutex(true, MutexName, out createdNew);

            if (!createdNew)
            {
                // Another instance is already collecting — send our file to it and exit
                try
                {
                    using var client = new System.IO.Pipes.NamedPipeClientStream(".", PipeName, System.IO.Pipes.PipeDirection.Out);
                    client.Connect(500);
                    using var writer = new StreamWriter(client);
                    writer.WriteLine(path);
                }
                catch { }
                return;
            }

            // First instance — collect files sent by other instances for 300ms, then open one window
            var files = new System.Collections.Generic.List<string> { path };

            var collectTask = System.Threading.Tasks.Task.Run(() =>
            {
                var deadline = DateTime.Now.AddMilliseconds(300);
                while (DateTime.Now < deadline)
                {
                    try
                    {
                        using var server = new System.IO.Pipes.NamedPipeServerStream(PipeName,
                            System.IO.Pipes.PipeDirection.In, 10,
                            System.IO.Pipes.PipeTransmissionMode.Byte,
                            System.IO.Pipes.PipeOptions.Asynchronous);
                        int waitMs = (int)(deadline - DateTime.Now).TotalMilliseconds;
                        if (waitMs <= 0) break;
                        if (server.WaitForConnectionAsync().Wait(waitMs))
                        {
                            using var reader = new StreamReader(server);
                            var line = reader.ReadLine();
                            if (line != null && (File.Exists(line) || Directory.Exists(line)))
                                lock (files) files.Add(line);
                        }
                    }
                    catch { break; }
                }
            });

            collectTask.Wait();
            mutex.ReleaseMutex();

            Application.Run(new Form1(files.ToArray()));
        }
    }
}