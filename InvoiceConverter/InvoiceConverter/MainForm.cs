using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace test
{
    public partial class MainForm : Form
    {
        private static BaseColor white = new BaseColor(255, 255, 255);
        private static BaseColor grey = new BaseColor(110, 110, 110);

        private static iTextSharp.text.Font ftSmerWhite = FontFactory.GetFont("Helvetica Neue", 9, iTextSharp.text.Font.NORMAL, BaseColor.WHITE);
        private static iTextSharp.text.Font ftSmer = FontFactory.GetFont("Helvetica Neue", 9, iTextSharp.text.Font.NORMAL, BaseColor.BLACK);
        private static iTextSharp.text.Font ftSm = FontFactory.GetFont("Helvetica Neue", 10, iTextSharp.text.Font.NORMAL, BaseColor.BLACK);
        private static iTextSharp.text.Font ftSmBold = FontFactory.GetFont("Helvetica Neue", 10, iTextSharp.text.Font.BOLD, BaseColor.BLACK);
        private static iTextSharp.text.Font ftNormal = FontFactory.GetFont("Helvetica Neue", 12, iTextSharp.text.Font.NORMAL, BaseColor.BLACK);
        private static iTextSharp.text.Font ftNormalWhite = FontFactory.GetFont("Helvetica Neue", 12, iTextSharp.text.Font.NORMAL, BaseColor.WHITE);
        private static iTextSharp.text.Font ftBig = FontFactory.GetFont("Helvetica Neue", 14, iTextSharp.text.Font.BOLD, BaseColor.BLACK);
        private static iTextSharp.text.Font ftBiger = FontFactory.GetFont("Helvetica Neue", 16, iTextSharp.text.Font.BOLD, BaseColor.BLACK);
        private static iTextSharp.text.Font ftSign = FontFactory.GetFont("Helvetica Neue", 30, iTextSharp.text.Font.BOLD, BaseColor.BLACK);
        private static iTextSharp.text.Font ftSmWhite = FontFactory.GetFont("Helvetica Neue", 10, iTextSharp.text.Font.NORMAL, BaseColor.WHITE);
        private static iTextSharp.text.Font ftBgWhite = FontFactory.GetFont("Helvetica Neue", 28, iTextSharp.text.Font.BOLD, BaseColor.WHITE);

        private static int side_margin = 40;
        private static int top_margin = 30;

        private static string logo_path = @"cloudnet_logo.png";
        public MainForm()
        {
            InitializeComponent();
        }

        private string Convert(string oriFilePath, string newFilePath)
        {
            //Step 1: Create a Docuement-Object
            Document document = new Document();
            // we create a reader for the document
            PdfReader reader = new PdfReader(oriFilePath);

            try
            {
                //Step 2: we create a writer that listens to the document
                PdfWriter writer = PdfWriter.GetInstance(document, new FileStream(newFilePath, FileMode.Create));

                //Step 3: Open the document
                document.Open();

                PdfContentByte cb = writer.DirectContent;



                for (int pageNumber = 1; pageNumber < reader.NumberOfPages + 1; pageNumber++)
                {
                    var pageSize = reader.GetPageSizeWithRotation(1);
                    document.SetPageSize(pageSize);
                    document.NewPage();

                    //get page
                    PdfImportedPage page = writer.GetImportedPage(reader, pageNumber);

                    //add page
                    int rotation = reader.GetPageRotation(pageNumber);
                    if (rotation == 90 || rotation == 270)
                    {
                        cb.AddTemplate(page, 0, -1f, 1f, 0, 0, reader.GetPageSizeWithRotation(pageNumber).Height);
                    }
                    else
                    {
                        cb.AddTemplate(page, 1f, 0, 0, 1f, 0, 0);
                    }

                    //replace header
                    if (Contains(reader, pageNumber, "GST Registration Number"))
                    {
                        ReplaceLogo(cb, pageSize);
                    }

                    //modify pay slip
                    if (Contains(reader, pageNumber, "Payment Slip"))
                    {
                        ModifyPaySlip(cb, pageSize);
                    }
                }
            }
            catch (Exception e)
            {
                return (e.Message);
            }
            finally
            {
                document.Close();
                reader.Close();
            }

            return "OK";
        }

        private bool Contains(PdfReader reader, int page, string str)
        {
            var content = PdfTextExtractor.GetTextFromPage(reader, page);
            return content.Contains(str);
        }

        private void ReplaceLogo(PdfContentByte context, iTextSharp.text.Rectangle pageSize)
        {
            //cover original header
            context.Rectangle(0, pageSize.Height - 120, pageSize.Width, 120);
            context.SetColorStroke(white);
            context.SetColorFill(white);
            context.FillStroke();

            //logo
            var logo = iTextSharp.text.Image.GetInstance(logo_path);
            logo.ScaleToFit(211, 60);
            logo.SetAbsolutePosition(side_margin - 20, pageSize.Height - 50 - top_margin);
            context.AddImage(logo);

            //caption
            var statementInfo = new Phrase("Statement / Tax Invoice", ftNormal);
            ColumnText.ShowTextAligned(context, Element.ALIGN_LEFT, statementInfo, side_margin, pageSize.Height - 65 - top_margin, 0);

            var gstInfo = new Phrase("GST Registration Number: 111-597-251", ftNormal);
            ColumnText.ShowTextAligned(context, Element.ALIGN_LEFT, gstInfo, side_margin, pageSize.Height - 80 - top_margin, 0);


            //contact info
            var contactL1 = new Phrase("CloudNet Services Ltd.", ftSm);
            var contactL2 = new Phrase("Level 5, 396 Queen Street, Auckland 1010, New Zealand", ftSm);
            var contactL3 = new Phrase("Tel: 0800 00 1745", ftSm);
            var contactL4 = new Phrase("http://www.cloudnetnz.com", ftSm);
            ColumnText.ShowTextAligned(context, Element.ALIGN_RIGHT, contactL1, pageSize.Width - side_margin, pageSize.Height - 40 - top_margin, 0);
            ColumnText.ShowTextAligned(context, Element.ALIGN_RIGHT, contactL2, pageSize.Width - side_margin, pageSize.Height - 52 - top_margin, 0);
            ColumnText.ShowTextAligned(context, Element.ALIGN_RIGHT, contactL3, pageSize.Width - side_margin, pageSize.Height - 64 - top_margin, 0);
            ColumnText.ShowTextAligned(context, Element.ALIGN_RIGHT, contactL4, pageSize.Width - side_margin, pageSize.Height - 76 - top_margin, 0);

            //hr
            //context.Rectangle(side_margin, pageSize.Height - 90 - top_margin, pageSize.Width - side_margin * 2, 3);
            //context.SetColorStroke(lightGreen);
            //context.SetColorFill(lightGreen);
            //context.FillStroke();
        }

        private void ModifyPaySlip(PdfContentByte context, iTextSharp.text.Rectangle pageSize)
        {
            //cover original logo
            context.Rectangle(30, 105, 180, 110);
            context.SetColorStroke(white);
            context.SetColorFill(white);
            context.FillStroke();

            //payment slip
            var slipInfo = new Phrase("Payment Slip", ftSmBold);
            ColumnText.ShowTextAligned(context, Element.ALIGN_LEFT, slipInfo, 120, 210, 0);

            //logo
            var logo = iTextSharp.text.Image.GetInstance(logo_path);
            logo.ScaleToFit(157, 45);
            logo.SetAbsolutePosition(20, 170);
            context.AddImage(logo);

            //payment
            var paybyInfo = new Phrase("Paying By Direct Credit", ftSmer);
            ColumnText.ShowTextAligned(context, Element.ALIGN_LEFT, paybyInfo, 35, 150, 0);

            var bankInfo = new Phrase("Bank: ASB", ftSmer);
            ColumnText.ShowTextAligned(context, Element.ALIGN_LEFT, bankInfo, 35, 140, 0);

            var accountInfo = new Phrase("Name of Account: CloudNet Services Limited", ftSmer);
            ColumnText.ShowTextAligned(context, Element.ALIGN_LEFT, accountInfo, 35, 125, 0);

            var accountNumInfo = new Phrase("Account Number: 12-3029-0439029-00", ftSmer);
            ColumnText.ShowTextAligned(context, Element.ALIGN_LEFT, accountNumInfo, 35, 115, 0);


            //cheque part
            context.Rectangle(35, 23, 265, 75);
            context.SetColorStroke(grey);
            context.SetColorFill(grey);
            context.FillStroke();

            var chequeInfo = new Phrase("Paying by cheques", ftSmerWhite);
            ColumnText.ShowTextAligned(context, Element.ALIGN_LEFT, chequeInfo, 35, 90, 0);

            var companyInfo = new Phrase("Please make cheques payable to CloudNet Services Ltd. and", ftSmerWhite);
            ColumnText.ShowTextAligned(context, Element.ALIGN_LEFT, companyInfo, 35, 80, 0);

            var companyInfo2 = new Phrase("write your Name and Phone Number on the back of your cheque.", ftSmerWhite);
            ColumnText.ShowTextAligned(context, Element.ALIGN_LEFT, companyInfo2, 35, 70, 0);

            var postInfo = new Phrase("Please post it with this payment slip to", ftSmerWhite);
            ColumnText.ShowTextAligned(context, Element.ALIGN_LEFT, postInfo, 35, 50, 0);
            var postInfo2 = new Phrase("CloudNet Services Limited", ftSmerWhite);
            ColumnText.ShowTextAligned(context, Element.ALIGN_LEFT, postInfo2, 35, 40, 0);
            var postInfo3 = new Phrase("Level 5, 396 Queen Street, Auckland 1010, New Zealand", ftSmerWhite);
            ColumnText.ShowTextAligned(context, Element.ALIGN_LEFT, postInfo3, 35, 30, 0);
        }

        private void btnConvert_Click(object sender, EventArgs e)
        {
            //string oldFile = @"C:\Works\CloudNet\InvoiceConverter\img\invoice_103186.pdf";
            //string newFile = oldFile.Insert(oldFile.LastIndexOf('.'), "_new");
            btnConvert.Enabled = false;
            toolStripStatusLblStatus.Text = "Converting ...";
            var oriFiles = GetOriginalFilePaths(txtFolder.Text);
            if (oriFiles.Count == 0)
            {
                MessageBox.Show("Plesae select a folder of invoice pdf files");
                btnConvert.Enabled = true;
                return;
            }
            else
            {
                toolStripProgressBar.Maximum = oriFiles.Count;
                toolStripProgressBar.Value = 0;
                txtLog.Text += string.Format("found {0} invoices\r\n", oriFiles.Count);
                foreach (var oriFile in oriFiles)
                {
                    var newFile = oriFile.Insert(oriFile.LastIndexOf("."), "_CloudNet");
                    var msg = Convert(oriFile, newFile);
                    txtLog.Text += string.Format("{0} ...... {1}\r\n", oriFile.Substring(oriFile.LastIndexOf('\\') + 1), msg);
                    toolStripProgressBar.Value++;
                }
                toolStripStatusLblStatus.Text = "Done";
                txtLog.Text += "Done";
            }

            btnConvert.Enabled = true;
        }

        private void btnFolder_Click(object sender, EventArgs e)
        {
            var result = folderBrowserDialog.ShowDialog();
            txtFolder.Text = folderBrowserDialog.SelectedPath;
        }

        private IList<string> GetOriginalFilePaths(string folder)
        {
            var result = new List<string>();
            if (string.IsNullOrEmpty(txtFolder.Text))
            {
                return result;
            }
            //check folder
            if (!Directory.Exists(folder))
            {
                return result;
            }

            var files = Directory.GetFiles(folder);

            //load file
            foreach (var file in files)
            {
                var fileLower = file.ToLower();
                fileLower = fileLower.Substring(fileLower.LastIndexOf('\\') + 1);
                if (fileLower.StartsWith("invoice") && fileLower.EndsWith("pdf"))
                {
                    result.Add(file);
                }
            }

            return result;
        }

        private void MainForm_Load(object sender, EventArgs e)
        {

        }
    }
}