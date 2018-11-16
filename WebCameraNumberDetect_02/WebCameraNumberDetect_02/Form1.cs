using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Windows.Forms;
using AForge.Video.DirectShow;
using dshow;
using Tesseract;

namespace WebCameraNumberDetect_02
{
    public partial class Form1 : Form
    {
        string device;
        FilterCollection filters;
        private VideoCaptureDevice FinalVideo = null;
        private Bitmap video;
        private Bitmap videoWithRect;
        
        //private AForge.Video.FFMPEG.VideoFileWriter FileWriter = new AForge.Video.FFMPEG.VideoFileWriter();
        int P = 0;
        UserRect rect;
        TesseractEngine Tesseractengine;
        List<ExcelValues> dict;
        double confidence;

        public Form1()
        {
            InitializeComponent();
            rect = new UserRect(new Rectangle(10, 10, 100, 100));
            rect.SetPictureBox(this.pictureBox1);
            

            string[] files = Directory.GetFiles(@"tessdata", "*.traineddata");
            for (int i = 0; i < files.Length; i++)
            {
                comboBox1.Items.Add(Path.GetFileNameWithoutExtension(files[i]));
            }
            comboBox1.SelectedIndex = 0;
            TesseractEngineLoad(comboBox1.Text);

            string s = Path.GetFileNameWithoutExtension(files[0]);

            try
            {
                filters = new FilterCollection(FilterCategory.VideoInputDevice);

                if (filters.Count == 0)
                    throw new ApplicationException();
                foreach (Filter item in filters)
                    comboBox2.Items.Add(item.Name);
                comboBox2.SelectedIndex = 0;
                

            }
            catch (ApplicationException)
            {
                MessageBox.Show("Не найден вебкамера");
            }
            dict = new List<ExcelValues>();
        }

        void TesseractEngineLoad(string font)
        {
            if (!string.IsNullOrEmpty(font))
            {
                if (Tesseractengine != null)
                {
                    Tesseractengine.Dispose();
                    Tesseractengine = null;
                }
                Tesseractengine = new TesseractEngine(@"tessdata", font, EngineMode.TesseractOnly);
                Tesseractengine.SetVariable("tessedit_char_whitelist", "0123456789:.");
            }
            
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            FinalVideo = new VideoCaptureDevice();
            FinalVideo.Source = device;// = captureDevice.VideoDevice;
            FinalVideo.NewFrame += new AForge.Video.NewFrameEventHandler(FinalVideo_NewFrame);
            FinalVideo.Start();
            timer1.Enabled = true;
            timer1.Start();
            dateTimePicker1.Value = new DateTime(2000, 01, 01, 0, 0, 1);
        }

        void ChartUpdate(double value)
        {
            chart1.Series["Series1"].Points.AddY(value);
        }

        void FinalVideo_NewFrame(object sender, AForge.Video.NewFrameEventArgs eventArgs)
        {
            if (false)
            {
                video = (Bitmap)eventArgs.Frame.Clone();
                pictureBox1.Image = (Bitmap)eventArgs.Frame.Clone();
                //AVIwriter.Quality = 0;
                //FileWriter.WriteVideoFrame(video);
                //AVIwriter.AddFrame(video);
            }
            else
            {
                video = (Bitmap)eventArgs.Frame.Clone();

                if (!checkBox1.Checked && !checkBox2.Checked)
                {
                    videoWithRect = new Bitmap(rect.rect.Width, rect.rect.Height);
                    using (Graphics g = Graphics.FromImage(videoWithRect))
                        g.DrawImage((Bitmap)eventArgs.Frame.Clone(), new Rectangle(0, 0, videoWithRect.Width, videoWithRect.Height), rect.rect, GraphicsUnit.Pixel);
                }
                else if (checkBox1.Checked && checkBox2.Checked)
                {
                    var videoWithRect1 = new Bitmap(rect.rect.Width, rect.rect.Height);
                    using (Graphics g = Graphics.FromImage(videoWithRect1))
                        g.DrawImage((Bitmap)eventArgs.Frame.Clone(), new Rectangle(0, 0, videoWithRect1.Width, videoWithRect1.Height), rect.rect, GraphicsUnit.Pixel);

                    videoWithRect1 = Transform((Bitmap)videoWithRect1.Clone());
                    videoWithRect = BlackAndWhite((Bitmap)videoWithRect1.Clone());
                }
                else if(checkBox2.Checked)
                {
                    var videoWithRect1 = new Bitmap(rect.rect.Width, rect.rect.Height);
                    using (Graphics g = Graphics.FromImage(videoWithRect1))
                        g.DrawImage((Bitmap)eventArgs.Frame.Clone(), new Rectangle(0, 0, videoWithRect1.Width, videoWithRect1.Height), rect.rect, GraphicsUnit.Pixel);
                    videoWithRect = Transform((Bitmap)videoWithRect1.Clone());
                }
                else if(checkBox1.Checked)
                {
                    var videoWithRect1 = new Bitmap(rect.rect.Width, rect.rect.Height);
                    using (Graphics g = Graphics.FromImage(videoWithRect1))
                        g.DrawImage((Bitmap)eventArgs.Frame.Clone(), new Rectangle(0, 0, videoWithRect1.Width, videoWithRect1.Height), rect.rect, GraphicsUnit.Pixel);
                    videoWithRect = BlackAndWhite((Bitmap)videoWithRect1.Clone());
                }

                pictureBox1.Image = video;
                pictureBox2.Image = videoWithRect;
            }
        }

        void OCR(Bitmap image)
        {
            try
            {
                MemoryStream byteStream = new MemoryStream();
                image.Save(byteStream, System.Drawing.Imaging.ImageFormat.Tiff);

                label1.Invoke(new MethodInvoker(delegate () { label1.Text = ""; }));

                using (var img = Pix.LoadTiffFromMemory(byteStream.ToArray()))
                {
                    using (var page = Tesseractengine.Process(img, PageSegMode.SingleWord))
                    {
                        confidence = page.GetMeanConfidence() * 100;
                        label6.Invoke(new MethodInvoker(delegate () { label6.Text = (confidence / 100).ToString("P"); }));
                        var text = page.GetText().Replace("\r", "").Replace("\n", "").Replace(" ", "");
                        label1.Invoke(new MethodInvoker(delegate () { label1.Text += text; }));
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Unexpected Error: " + e.Message);
                Console.WriteLine("Details: ");
                Console.WriteLine(e.ToString());
            }
        }

        public Bitmap Transform(Bitmap source)
        {
            //create a blank bitmap the same size as original
            Bitmap newBitmap = new Bitmap(source.Width, source.Height);

            //get a graphics object from the new image
            Graphics g = Graphics.FromImage(newBitmap);

            // create the negative color matrix
            ColorMatrix colorMatrix = new ColorMatrix(new float[][]
            {
        new float[] {-1, 0, 0, 0, 0},
        new float[] {0, -1, 0, 0, 0},
        new float[] {0, 0, -1, 0, 0},
        new float[] {0, 0, 0, 1, 0},
        new float[] {1, 1, 1, 0, 1}
            });

            // create some image attributes
            ImageAttributes attributes = new ImageAttributes();

            attributes.SetColorMatrix(colorMatrix);

            g.DrawImage(source, new Rectangle(0, 0, source.Width, source.Height),
                        0, 0, source.Width, source.Height, GraphicsUnit.Pixel, attributes);

            //dispose the Graphics object
            g.Dispose();

            return newBitmap;
        }

        private Bitmap BlackAndWhite(Bitmap bmpImg)
        {
            Color color = new Color();
            try
            {
                for (int j = 0; j < bmpImg.Height; j++)
                {
                    for (int i = 0; i < bmpImg.Width; i++)
                    {
                        color = bmpImg.GetPixel(i, j);
                        int K = (color.R + color.G + color.B) / 3;
                        bmpImg.SetPixel(i, j, K <= P ? Color.Black : Color.White);
                    }
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return bmpImg;
        }

        private Bitmap BlackAndWhite1(Bitmap bmpImg)
        {
            //BitmapData bmData = bmpImg.LockBits(new Rectangle(0, 0, bmpImg.Width, bmpImg.Height), ImageLockMode.ReadWrite, PixelFormat.Format24bppRgb);
            PixelUtil pixelUtil = new PixelUtil(bmpImg);
            pixelUtil.LockBits();
            Color color = new Color();
            try
            {
                for (int j = 0; j < bmpImg.Height; j++)
                {
                    for (int i = 0; i < bmpImg.Width; i++)
                    {
                        color = pixelUtil.GetPixel(i, j);
                        int K = (color.R + color.G + color.B) / 3;
                        pixelUtil.SetPixel(i, j, K <= P ? Color.Black : Color.White);
                    }
                }
            }

            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            pixelUtil.UnlockBits();
            return bmpImg;
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            FinalVideo.Stop();
            timer1.Stop();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (videoWithRect == null)
                return;
            if (Tesseractengine == null)
                return;
            OCR((Bitmap)videoWithRect);

            if (buttonStart)
            {
                var time = DateTime.Now;
                string timeString = string.Format("{0}:{1}:{2}", time.Hour, time.Minute, time.Second);
                label3.Invoke(new MethodInvoker(delegate () { label3.Text = timeString; }));
                if (confidence >= (double)numericUpDown1.Value)
                {
                    double y = 0;
                    if (GetDouble(label1.Text, out y))
                    {
                        textBox1.AppendText(timeString + " - " + y + "\r\n");
                        ChartUpdate(y);
                        dict.Add(new ExcelValues() { time = new TimeSpan(time.Hour, time.Minute, time.Second), value = y });
                    }
                    else
                    {
                        string s = label1.Text.Replace(':', '.');
                        if (GetDouble(s, out y))
                        {
                            textBox1.AppendText(timeString + " - " + y + "\r\n");
                            ChartUpdate(y);
                            dict.Add(new ExcelValues() { time = new TimeSpan(time.Hour, time.Minute, time.Second), value = y });
                        }
                    }
                }
            }
        }

        public static bool GetDouble(string value, out double defaultValue)
        {
            double result;

            //Try parsing in the current culture
            if (!double.TryParse(value, System.Globalization.NumberStyles.Any, CultureInfo.CurrentCulture, out result) &&
                //Then try in US english
                !double.TryParse(value, System.Globalization.NumberStyles.Any, CultureInfo.GetCultureInfo("en-US"), out result) &&
                //Then in neutral language
                !double.TryParse(value, System.Globalization.NumberStyles.Any, CultureInfo.InvariantCulture, out result))
            {
                defaultValue = 0;
                return false;
            }
            defaultValue = result;
            return true;
        }

        private void trackBar1_Scroll(object sender, EventArgs e)
        {
            P = trackBar1.Value;
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            TimeSpan t = new TimeSpan(dateTimePicker1.Value.Hour, dateTimePicker1.Value.Minute, dateTimePicker1.Value.Second);
            if ((int)t.Ticks > 0)
            {
                timer1.Interval = (int)t.TotalMilliseconds;
            }
        }

        bool buttonStart = false;
        private void button1_Click(object sender, EventArgs e)
        {
            if (buttonStart)
            {
                buttonStart = false;
                button1.Text = "Start";
                button1.BackColor = Color.Green;
                dateTimePicker1.Enabled = true;
            }
            else
            {
                buttonStart = true;
                label3.Text = "";
                button1.Text = "Stop";
                button1.BackColor = Color.Red;
                dateTimePicker1.Enabled = false;
                textBox1.Clear();
                chart1.Series["Series1"].Points.Clear();
                dict.Clear();
            }
        }

        private delegate void CloseDelegate();
        private void button2_Click(object sender, EventArgs e)
        {
            if (dict.Count != 0)
            {
                tableLayoutPanel1.Enabled = false;
                PleaseWait pw = new PleaseWait();
                CloseDelegate cDel = new CloseDelegate(pw.Close);
                System.Threading.Thread th = new System.Threading.Thread(pw.Show);
                th.Start();

                //BeginInvoke((Action)(() => pw.ShowDialog(this)));
                ExcelHelper h = new ExcelHelper();
                h.Create(dict);
                this.Invoke(cDel);
                pw = null;
                tableLayoutPanel1.Enabled = true;
                MessageBox.Show("Done \r\n informations.xls");
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://youtu.be/FZRbLmYi2VA");
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                trackBar1.Enabled = true;
            }
            else
            {
                trackBar1.Enabled = false;
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            TesseractEngineLoad(comboBox1.Text);
        }

        private void buttonCameraSetting_Click(object sender, EventArgs e)
        {
            if (FinalVideo != null)
                FinalVideo.DisplayPropertyPage(this.Handle);
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            foreach (Filter item in filters)
            {
                if (item.Name.Contains(comboBox2.Text))
                {
                    device = item.MonikerString;
                    if (FinalVideo != null)
                    {
                        FinalVideo.Stop();
                        FinalVideo.Source = device;// = captureDevice.VideoDevice;
                        FinalVideo.Start();
                    }
                }
            }
        }
    }
    public class ExcelValues
    {
        public TimeSpan time;
        public double value;
    }
}
