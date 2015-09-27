using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Drawing2D;
using System.Runtime.InteropServices;
using System.IO;
using System.Globalization;
using System.Windows.Forms;
using System.Threading;
using System.Diagnostics;

namespace OCRSample
{
    //*****************************************************************************************************
    
    class PersistenceSingleton {

        private static readonly PersistenceSingleton instance = new PersistenceSingleton();

        private PersistenceSingleton() { }

        public static PersistenceSingleton getInstance
        {
            get
            {
                return instance;
            }
        }

        public Point[] zonesbar = null;
    }
    
    class Skype {

        private static PersistenceSingleton singleton = PersistenceSingleton.getInstance;
        private delegate void ScannerReadyEventHandler(Point[] points);
        private static event ScannerReadyEventHandler startHandler;

        //Trace.WriteLine(obj);
        //Debug.WriteLine(zones.Length);
        //Console.WriteLine(zones.Length);

        [STAThread]
        public static void Main() {
            RunDetectZonesBar();
            Application.Run();
        }

        private static void RunDetectZonesBar() {
            Skype.startHandler = scannerReady;
            IntPtr hwnd = Wim32API.FindWindow("tSkMainForm", null);
            IntPtr hwndChild = Wim32API.FindWindowEx(hwnd, IntPtr.Zero, "TAppToolbarControl", null);
            Worker.ZonesEventHandler handler = DelegateZones;
            Worker workerObject = new Worker(hwnd, hwndChild, handler);
            Thread workerThread = new Thread(workerObject.initDetectZones);
            workerThread.Start();
        }

        private static void DelegateZones(Rectangle panel, Rectangle[] zones) {
            if (panel != null && zones != null && zones.Length>0) {
                Point[] points = new Point[zones.Length];
                for (int i = 0; i < zones.Length; i++) {
                    int middle = zones[i].Height / 2;
                    int X = zones[i].X + panel.X + middle;
                    int Y = zones[i].Y + panel.Y + middle;
                    points[i] = new Point(X, Y);
                }
                if (startHandler != null) startHandler.Invoke(points);
            }
        }

        private static void cleanAndSearch(string text) {
            if (singleton.zonesbar != null && singleton.zonesbar.Length == 4) {
                int index = 3;
                Wim32API.SetCursorPos(singleton.zonesbar[index].X, singleton.zonesbar[index].Y);
                Wim32API.mouseClick(singleton.zonesbar[index].X, singleton.zonesbar[index].Y, true);
                Thread.Sleep(100);

                index = 0;
                Wim32API.SetCursorPos(singleton.zonesbar[index].X, singleton.zonesbar[index].Y);
                Wim32API.mouseClick(singleton.zonesbar[index].X, singleton.zonesbar[index].Y, true);
                Thread.Sleep(100);

                index = 3;
                Wim32API.SetCursorPos(singleton.zonesbar[index].X, singleton.zonesbar[index].Y);
                Wim32API.mouseClick(singleton.zonesbar[index].X, singleton.zonesbar[index].Y, true);
                Thread.Sleep(100);

                Wim32API.SendText(text);
            }
        }

        public static Bitmap toBlack(Bitmap bmp) { 
            Bitmap bt = (Bitmap)bmp.Clone();
            for (int i = 0; i < bt.Width; i++) {
                for (int j = 0; j < bt.Height; j++) { 
                    Color c = bt.GetPixel(i, j);
                    if (c.R != 255 && c.G != 255 && c.B != 255) {
                        c = Color.FromArgb(0, 0, 0);
                        bt.SetPixel(i, j, c);
                    }
                }
            }
            return bt;
        }

        private static object[] getCurrentSearch()
        {
            object[] currentItems = null;
            Bitmap[] item = OCRTools.getCurrentItem();
            if (item != null && item.Length == 2) {
                OCRTools.saveBitmap(item[0], "capture_photo", true);
                OCRTools.saveBitmap(item[1], "capture_name", true);
                string name = OCRDecoder.getTextBitmap("capture_name");
                OCRTools.deleteBitmap("capture_name", true);
                if (name != null && name.Length > 0) {
                    currentItems = new object[2];
                    currentItems[0] = item[0];
                    currentItems[1] = name;
                }
            }
            return currentItems;
        }

        private static void scannerReady(Point[] points){
            singleton.zonesbar = points;
            cleanAndSearch("nicolle-xekri@hotmail.com");
            Thread.Sleep(2 * 1000);
            object[] firstSearch = getCurrentSearch();
            cleanAndSearch("ivan.sanchez@nullcode.com.ar");
            Thread.Sleep(2 * 1000);
            object[] secondSearch = getCurrentSearch();

            MessageBox.Show("Found: nicolle-xekri@hotmail.com -> " + firstSearch[1] + " & ivan.sanchez@nullcode.com.ar -> " + secondSearch[1]);

            secondSearch = null;
        }

        
        /*private static Bitmap CreateLogo_quitar(string subdomain) {
            Bitmap objBmpImage = new Bitmap(1, 1);
            int intWidth = 0;
            int intHeight = 0;
            Font objFont = new Font("Tahoma", 12, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            Graphics objGraphics = Graphics.FromImage(objBmpImage);
            intWidth = (int)objGraphics.MeasureString(subdomain, objFont).Width;
            intHeight = (int)objGraphics.MeasureString(subdomain, objFont).Height;
            objBmpImage = new Bitmap(objBmpImage, new Size(intWidth, intHeight));
            objGraphics = Graphics.FromImage(objBmpImage);
            objGraphics.Clear(Color.Transparent);
            objGraphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            //objGraphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;
            objGraphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.SingleBitPerPixelGridFit;
            //objGraphics.DrawString(subdomain, objFont, new SolidBrush(Color.FromArgb(255, 255, 255)), 0, 0);
            objGraphics.DrawString(subdomain, objFont, new SolidBrush(Color.Black), 0, 0);
            objGraphics.Flush();
            return (objBmpImage);
        }*/

    }

    class Wim32API {
        
        //*****************************************************************************************************

        [DllImport("user32.dll")]
        public static extern bool SetCursorPos(int X, int Y);

        [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        public static extern void mouse_event(long dwFlags, long dx, long dy, long cButtons, long dwExtraInfo);

        public const int MOUSEEVENTF_LEFTDOWN = 0x02;
        public const int MOUSEEVENTF_LEFTUP = 0x04;
        public const int MOUSEEVENTF_RIGHTDOWN = 0x08;
        public const int MOUSEEVENTF_RIGHTUP = 0x10;

        public static void mouseClick(long dx, long dy, bool isLeft) {
            if (isLeft) {
                mouse_event(MOUSEEVENTF_LEFTDOWN, dx, dy, 0, 0);
                mouse_event(MOUSEEVENTF_LEFTUP, dx, dy, 0, 0);
            }
            else {
                mouse_event(MOUSEEVENTF_RIGHTDOWN, dx, dy, 0, 0);
                mouse_event(MOUSEEVENTF_RIGHTUP, dx, dy, 0, 0);
            }
        }

        //*****************************************************************************************************

        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

        [DllImport("user32.dll")]
        public static extern IntPtr GetDC(IntPtr WindowHandle);

        [DllImport("user32.dll")]
        public static extern void ReleaseDC(IntPtr WindowHandle, IntPtr DC);

        [DllImport("user32.dll")]
        public static extern IntPtr GetWindowRect(IntPtr WindowHandle, ref Rect rect);

        [DllImport("User32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool PrintWindow(IntPtr hwnd, IntPtr hDC, uint nFlags);

        [StructLayout(LayoutKind.Sequential)]
        public struct Rect
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        //*****************************************************************************************************

        public static void SendText(string text) {
            System.Windows.Forms.SendKeys.SendWait(text);
        }

        [DllImport("User32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int uMsg, int wParam, int lParam);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        public static extern IntPtr SetFocus(IntPtr hwnd);

        [DllImport("user32", CharSet = CharSet.Auto, SetLastError = true, ExactSpelling = true)]
        public static extern IntPtr GetWindow(IntPtr hwnd, int wFlag);

        //*****************************************************************************************************

    }

    class Worker {

        public delegate void ZonesEventHandler(Rectangle panel, Rectangle[] zones);
        private event ZonesEventHandler detectedHandler;

        private IntPtr hwndChild = IntPtr.Zero;
        private IntPtr hwnd = IntPtr.Zero;

        public Worker(IntPtr hwnd, IntPtr hwndChild, ZonesEventHandler detectedHandler) {
            this.hwnd = hwnd;
            this.hwndChild = hwndChild;
            this.detectedHandler = detectedHandler;
        }

        private static Point detectTopLeft(Bitmap bmpInitial, bool detectLigth)
        {
            Point topLeft = new Point(-1, -1);
            int iconSize = 36;
            for (int y = 0; y < bmpInitial.Height; y++)
            {
                for (int x = 0; x < bmpInitial.Width; x++)
                {
                    Color c = bmpInitial.GetPixel(x, y);
                    int R = (int)c.R;
                    int G = (int)c.G;
                    int B = (int)c.B;
                    bool darkGray = false;
                    bool ligthGray = false;

                    if (detectLigth)
                    {
                        if (R == 229 && G == 229 && B == 229) ligthGray = true;
                    }
                    else
                    {
                        if (R == 110 && G == 110 && B == 110) darkGray = true;
                    }

                    if (topLeft.X == -1 && topLeft.Y == -1)
                    {
                        if (darkGray || ligthGray)
                        {
                            bool endColor = false;
                            if (x + iconSize <= bmpInitial.Width && y + iconSize <= bmpInitial.Height)
                            {
                                c = bmpInitial.GetPixel(x + (iconSize - 1), y + (iconSize - 1));
                                R = (int)c.R;
                                G = (int)c.G;
                                B = (int)c.B;
                                if (detectLigth)
                                {
                                    if (R == 229 && G == 229 && B == 229) endColor = true;
                                }
                                else
                                {
                                    if (R == 110 && G == 110 && B == 110) endColor = true;
                                }
                            }

                            if (endColor) topLeft = new Point(x, y);
                        }
                    }
                }
            }
            return topLeft;
        }

        private static Rectangle[] detectBarZones(Bitmap[] bmps)
        {
            List<Rectangle> rectList = new List<Rectangle>();
            if (bmps != null && bmps.Length > 0)
            {
                for (int i = 0; i < bmps.Length; i++)
                {
                    if (bmps[i] != null)
                    {
                        Point topLeftLigth = detectTopLeft(bmps[i], true);
                        if (topLeftLigth.X != -1 && topLeftLigth.Y != -1)
                        {
                            Rectangle zoneLigth = new Rectangle(topLeftLigth.X, topLeftLigth.Y, 36, 36);
                            if (!rectList.Contains(zoneLigth)) rectList.Add(zoneLigth);
                        }

                        Point topLeftDark = detectTopLeft(bmps[i], false);
                        if (topLeftDark.X != -1 && topLeftDark.Y != -1)
                        {
                            Rectangle zoneDark = new Rectangle(topLeftDark.X, topLeftDark.Y, 36, 36);
                            if (!rectList.Contains(zoneDark)) rectList.Add(zoneDark);
                        }
                    }
                }
            }
            if (rectList.Count == 4) return rectList.ToArray();
            return null;
        }

        public void initDetectZones() {

            Wim32API.Rect rectTabPanel = new Wim32API.Rect();
            Wim32API.GetWindowRect(hwndChild, ref rectTabPanel);
            Rectangle rectangleTabPanel = new Rectangle();
            rectangleTabPanel.X = rectTabPanel.Left;
            rectangleTabPanel.Y = rectTabPanel.Top;
            rectangleTabPanel.Width = rectTabPanel.Right - rectTabPanel.Left + 1;
            rectangleTabPanel.Height = rectTabPanel.Bottom - rectTabPanel.Top + 1;

            Wim32API.SetForegroundWindow(hwnd);

            Bitmap bmpInitial = OCRTools.getBitmap(hwndChild);
            bmpInitial = OCRTools.toGrayscale(bmpInitial);

            List<Bitmap> bmpList = new List<Bitmap>();
            bmpList.Add(bmpInitial);

            int Y = rectangleTabPanel.Y + (rectangleTabPanel.Height / 2);
            int X = rectangleTabPanel.X + rectangleTabPanel.Width;

            while (true) {
                Wim32API.SetCursorPos(X, Y);
                X -= 10;
                Bitmap bmp = OCRTools.getBitmap(hwndChild);
                bmp = OCRTools.toGrayscale(bmp);
                if (bmp != null && bmp != bmpInitial && !isBlackImage(bmp) && !containsArray(bmpList, bmp)) bmpList.Add(bmp);
                if (X < rectangleTabPanel.X) break;
                Thread.Sleep(100);
            }

            Bitmap[] capturedBitmaps = null;
            if (bmpList.Count > 0){
                capturedBitmaps = bmpList.ToArray();
                bmpList.Clear();
            }

            if(capturedBitmaps != null && capturedBitmaps.Length>0){
                Rectangle[] zonesBar = detectBarZones(capturedBitmaps);
                Array.Sort(zonesBar, delegate(Rectangle rect1, Rectangle rect2) {
                    int rsp = 0;
                    if (rect1.X < rect2.X) rsp = -1;
                    else if (rect1.X > rect2.X) rsp = 1;
                    return rsp;
                });
                capturedBitmaps = null;
                if (detectedHandler != null && zonesBar != null && zonesBar.Length > 0) detectedHandler.Invoke(rectangleTabPanel, zonesBar);
            }

        }

        private bool containsArray(List<Bitmap> bmpList, Bitmap bmp) {
            bool contains = false;
            if (bmpList.Count > 0) {
                for (int i = 0; i < bmpList.Count; i++) {
                    if (bmpList[i] == bmp) {
                        contains = true;
                        break;
                    }
                }
            }
            return contains;
        }

        private bool isBlackImage(Bitmap bmp){
            bool isBlackImage = false;
            Color c = bmp.GetPixel(0, 0);
            int R = (int)c.R;
            int G = (int)c.G;
            int B = (int)c.B;
            if (R != 255 && G != 255 && B != 255) isBlackImage = true;
            return isBlackImage;
        }
    }

    class OCRDecoder
    {
        public static string getTextBitmap(string nameImage)
        {
            string text = null;
            try {
                string file = Path.GetFullPath(nameImage + ".jpg");
                MODI.Document md = new MODI.Document();
                md.Create(file);
                md.OCR(MODI.MiLANGUAGES.miLANG_SPANISH, false, false);
                MODI.Image image = (MODI.Image)md.Images[0];
                MODI.Layout layout = image.Layout;

                text = image.Layout.Text.Trim().Replace("\r\n", "");
                md.Close(false);

            }catch(Exception e){
                text = null;
            }
            return text;
        }
    }

    class OCRTools {

        public static Bitmap toNegative(Bitmap sourceimage)
        {
            Bitmap bt = (Bitmap)sourceimage.Clone();
            Color c;
            for (int i = 0; i < bt.Width; i++)
            {
                for (int j = 0; j < bt.Height; j++)
                {
                    c = bt.GetPixel(i, j);
                    c = Color.FromArgb(255 - c.R, 255 - c.G, 255 - c.B);
                    bt.SetPixel(i, j, c);
                }
            }
            return bt;
        }

        public static Bitmap toGrayscale(Bitmap original)
        {
            Bitmap newBitmap = new Bitmap(original.Width, original.Height);
            for (int i = 0; i < original.Width; i++)
            {
                for (int j = 0; j < original.Height; j++)
                {
                    Color originalColor = original.GetPixel(i, j);
                    int grayScale = (int)((originalColor.R * .3) + (originalColor.G * .59) + (originalColor.B * .11));
                    Color newColor = Color.FromArgb(grayScale, grayScale, grayScale);
                    newBitmap.SetPixel(i, j, newColor);
                }
            }
            return newBitmap;
        }

        public static unsafe Bitmap toDilate(Bitmap SrcImage)
        {
            // Create Destination bitmap.
            Bitmap tempbmp = new Bitmap(SrcImage.Width, SrcImage.Height);

            // Take source bitmap data.
            BitmapData SrcData = SrcImage.LockBits(new Rectangle(0, 0, SrcImage.Width, SrcImage.Height), ImageLockMode.ReadOnly, PixelFormat.Format24bppRgb);

            // Take destination bitmap data.
            BitmapData DestData = tempbmp.LockBits(new Rectangle(0, 0, tempbmp.Width, tempbmp.Height), ImageLockMode.ReadWrite, PixelFormat.Format24bppRgb);

            // Element array to used to dilate.
            byte[,] sElement = new byte[5, 5] { 
                {0,0,1,0,0},
                {0,1,1,1,0},
                {1,1,1,1,1},
                {0,1,1,1,0},
                {0,0,1,0,0}
            };

            // Element array size.
            int size = 5;
            byte max, clrValue;
            int radius = size / 2;
            int ir, jr;

            {

                // Loop for Columns.
                for (int colm = radius; colm < DestData.Height - radius; colm++)
                {
                    // Initialise pointers to at row start.
                    byte* ptr = (byte*)SrcData.Scan0 + (colm * SrcData.Stride);
                    byte* dstPtr = (byte*)DestData.Scan0 + (colm * SrcData.Stride);

                    // Loop for Row item.
                    for (int row = radius; row < DestData.Width - radius; row++)
                    {
                        max = 0;
                        clrValue = 0;

                        // Loops for element array.
                        for (int eleColm = 0; eleColm < 5; eleColm++)
                        {
                            ir = eleColm - radius;
                            byte* tempPtr = (byte*)SrcData.Scan0 +
                                ((colm + ir) * SrcData.Stride);

                            for (int eleRow = 0; eleRow < 5; eleRow++)
                            {
                                jr = eleRow - radius;

                                // Get neightbour element color value.
                                clrValue = (byte)((tempPtr[row * 3 + jr] +
                                    tempPtr[row * 3 + jr + 1] + tempPtr[row * 3 + jr + 2]) / 3);

                                if (max < clrValue)
                                {
                                    if (sElement[eleColm, eleRow] != 0)
                                        max = clrValue;
                                }
                            }
                        }

                        dstPtr[0] = dstPtr[1] = dstPtr[2] = max;

                        ptr += 3;
                        dstPtr += 3;
                    }
                }
            }

            // Dispose all Bitmap data.
            SrcImage.UnlockBits(SrcData);
            tempbmp.UnlockBits(DestData);

            // return dilated bitmap.
            return tempbmp;
        }

        public static Bitmap getBitmap(IntPtr handle)
        {
            Bitmap bmp = null;
            try {
                Wim32API.Rect rect = new Wim32API.Rect();
                Wim32API.GetWindowRect(handle, ref rect);
                bmp = new Bitmap(rect.Right - rect.Left, rect.Bottom - rect.Top);
                Graphics memoryGraphics = Graphics.FromImage(bmp);
                IntPtr dc = memoryGraphics.GetHdc();
                bool success = Wim32API.PrintWindow(handle, dc, 0);
                memoryGraphics.ReleaseHdc(dc);
            }
            catch (Exception e) { bmp = null; }
            return bmp;
        }

        public static Bitmap loadBitmap(string fileImage)
        {
            Bitmap bmp = null;
            try {
                bmp = new Bitmap(fileImage);
            } catch (Exception e) { bmp = null; }
            return bmp;
        }

        public static void saveBitmap(Bitmap bmp, string fileImage, bool isJpeg)
        {
            try {
                if (bmp != null && isJpeg) bmp.Save(fileImage + ".jpg", ImageFormat.Jpeg);
                else if (bmp != null && !isJpeg) bmp.Save(fileImage + ".png", ImageFormat.Png);
            } catch (Exception e) { bmp = null; }
        }

        public static void deleteBitmap(string fileImage, bool isJpeg)
        {
            if (isJpeg) System.IO.File.Delete(fileImage + ".jpg");
            else System.IO.File.Delete(fileImage + ".png");
        }

        private static Bitmap resizeImage(Bitmap bmp, Size size)
        {
            Bitmap bmpScale = null;
            try {
                int sourceWidth = bmp.Width;
                int sourceHeight = bmp.Height;
                float nPercent = 0;
                float nPercentW = 0;
                float nPercentH = 0;
                nPercentW = ((float)size.Width / (float)sourceWidth);
                nPercentH = ((float)size.Height / (float)sourceHeight);
                if (nPercentH < nPercentW) nPercent = nPercentH;
                else nPercent = nPercentW;
                int destWidth = (int)(sourceWidth * nPercent);
                int destHeight = (int)(sourceHeight * nPercent);
                bmpScale = new Bitmap(destWidth, destHeight);
                Graphics g = Graphics.FromImage((Image)bmpScale);
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.DrawImage(bmp, 0, 0, destWidth, destHeight);
                g.Dispose();
            } catch (Exception e) { bmpScale = null; }
            return bmpScale;
        }

        private static Bitmap cutSkypeList(Bitmap list) {
            Bitmap bmp = null;
            try {
                if (list != null && list.Width > 0 && list.Height > 0) {
                    int headerPos = 37 + 25;
                    bmp = new Bitmap(list.Width, list.Height - headerPos);
                    Graphics g = Graphics.FromImage(bmp);
                    g.DrawImage(list, 0, 0 - headerPos, list.Width, list.Height);
                    g.Dispose();
                }
            }
            catch (Exception e) { bmp = null; }
            return bmp;
        }

        private static Bitmap captureSkypeList()
        {
            Bitmap bmp = null;
            IntPtr hwnd = Wim32API.FindWindow("tSkMainForm", null);
            if (hwnd != IntPtr.Zero) {
                IntPtr hwndChild = Wim32API.FindWindowEx((IntPtr)hwnd, IntPtr.Zero, "TConversationsControl", null);
                bmp = getBitmap(hwndChild);
                bmp = cutSkypeList(bmp);
            }
            return bmp;
        }

        private static Bitmap getFirstItemSkypeList(Bitmap list)
        {
            Bitmap bmp = null;
            try {
                if (list != null && list.Width > 0 && list.Height > 0) {
                    bmp = new Bitmap(list.Width - 2, 40 - 8);
                    Graphics g = Graphics.FromImage(bmp);
                    g.DrawImage(list, -1, -5, list.Width, list.Height);
                    g.Dispose();
                }
            }catch (Exception e) { bmp = null; }
            return bmp;
        }

        private static Bitmap extractPhoto(Bitmap list)
        {
            Bitmap bmp = null;
            try {
                if (list != null && list.Width > 0 && list.Height > 0) {
                    bmp = new Bitmap(32, 32);
                    Graphics g = Graphics.FromImage(bmp);
                    g.DrawImage(list, -8, 0, list.Width, list.Height);
                    g.Dispose();
                }
            }catch (Exception e) { bmp = null; }
            return bmp;
        }

        private static Bitmap extractItem(Bitmap list)
        {
            Bitmap item = null;
            try {
                if (list != null && list.Width > 0 && list.Height > 0) {
                    int x = 32 + 9;
                    Bitmap bmp = new Bitmap(list.Width - x, 32);
                    Graphics g = Graphics.FromImage(bmp);
                    g.DrawImage(list, 0 - x, 0, list.Width, list.Height);
                    g.Dispose();
                    item = OCRTools.divideItem(bmp);
                }
            }
            catch (Exception e) { item = null; }
            return item;
        }

        private static Bitmap divideItem(Bitmap list)
        {
            Bitmap bmp = null;
            try
            {
                if (list != null && list.Width > 0 && list.Height > 0) {
                    list = OCRTools.cleanImage(list);
                    bmp = new Bitmap(list.Width, 32 / 2);

                    Graphics g = Graphics.FromImage(bmp);
                    g.DrawImage(list, 0, 0, list.Width, list.Height);
                    g.Dispose();

                    bmp = OCRTools.toGrayscale(bmp);
                    saveBitmap(bmp, "name", false);
                    bmp = resizeImage(bmp, new Size(bmp.Width * 4, bmp.Height * 4));
                }
            }
            catch (Exception e) { bmp = null; }
            return bmp;
        }

        private static bool isEmptyBitmap(Bitmap bmp)
        { 
            bool result = true;
            try
            {
                if (bmp != null && bmp.Width > 0 && bmp.Height > 0) {
                    for(int y=0; y<bmp.Height; y++){
                        for(int x=0; x<bmp.Width; x++){
                            Color c = bmp.GetPixel(x, y);
                            int R = (int)c.R;
                            int G = (int)c.G;
                            int B = (int)c.B;
                            if (R != 255 && G != 255 && B != 255) {
                                result = false;
                                return result;
                            }
                        }
                    }
                }
            }
            catch (Exception e) { result = true; }
            return result;
        }

        public static Bitmap darkness(Bitmap bmp, int percent)
        {
            Bitmap bt = (Bitmap)bmp.Clone();
            for (int i = 0; i < bt.Width; i++)
            {
                for (int j = 0; j < bt.Height; j++)
                {
                    Color c = bt.GetPixel(i, j);
                    if (c.R != 255 && c.G != 255 && c.B != 255)
                    {
                        int level = (255 / 100) * percent;
                        int R = c.R, G = c.G, B = c.B;
                        if (c.R > level) R = c.R - level;
                        if (c.G > level) G = c.G - level;
                        if (c.B > level) B = c.B - level;
                        c = Color.FromArgb(R, G, B);
                        bt.SetPixel(i, j, c);
                    }
                }
            }
            return bt;
        }

        private static Bitmap cleanImage(Bitmap bmp)
        {
            Bitmap bt = (Bitmap)bmp.Clone();
            for (int i = 0; i < bt.Width; i++)
            {
                for (int j = 0; j < bt.Height; j++)
                {
                    Color c = bt.GetPixel(i, j);
                    int R = c.R, G = c.G, B = c.B;
                    bool removeColor = false;
                    if (G == 219 && B == 144) removeColor = true;
                    else if (G == 238 && B == 206) removeColor = true;
                    else if (G == 219 && B == 255) removeColor = true;
                    else if (G == 255 && B == 255) removeColor = true;
                    else if (G == 255 && B == 182) removeColor = true;
                    else if (R == 102 && G == 182 && B == 255) removeColor = true;
                    else if (G == 255 && B == 222) removeColor = true;
                    else if (G == 255) removeColor = true;
                    else if (G == 206) removeColor = true;
                    else if (R == 222 && G == 222 && B == 188) removeColor = true;
                    else if (R == 171 && G == 206 && B == 238) removeColor = true;
                    else if (R == 188 && G == 222 && B == 255) removeColor = true;

                    //else if (R == 188 && G == 206 && B == 206) removeColor = true;


                    if (removeColor)
                    {
                        c = Color.FromArgb(255, 255, 255);
                        bt.SetPixel(i, j, c);
                    }
                }
            }
            return bt;
        }

        public static Bitmap[] getCurrentItem() {
            Bitmap[] detected = null;

            Bitmap bmp = captureSkypeList();
            bmp = getFirstItemSkypeList(bmp);
            if (bmp != null) {
                Bitmap photo = extractPhoto(bmp);
                Bitmap item = extractItem(bmp);
                if (item != null && !isEmptyBitmap(photo) && !isEmptyBitmap(item)) {
                    detected = new Bitmap[2];
                    detected[0] = photo;
                    detected[1] = item;
                }
            }
            return detected;
        }

    }

}

