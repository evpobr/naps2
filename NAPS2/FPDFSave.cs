/*
    NAPS2 (Not Another PDF Scanner 2)
    http://sourceforge.net/projects/naps2/
    
    Copyright (C) 2009       Pavel Sorejs
    Copyright (C) 2012       Michael Adams
    Copyright (C) 2013       Peter De Leeuw
    Copyright (C) 2012-2013  Ben Olden-Cooligan

    This program is free software; you can redistribute it and/or
    modify it under the terms of the GNU General Public License
    as published by the Free Software Foundation; either version 2
    of the License, or (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.
*/

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using NAPS2.Lang.Resources;
using NAPS2.Pdf;
using NAPS2.Scan;

namespace NAPS2
{
    public partial class FPDFSave : Form
    {
        private readonly IPdfExporter pdfExporter;
        private readonly IErrorOutput errorOutput;

        public FPDFSave(IPdfExporter pdfExporter, IErrorOutput errorOutput)
        {
            InitializeComponent();
            this.pdfExporter = pdfExporter;
            this.errorOutput = errorOutput;
            Shown += FPDFSave_Shown;
        }

        public string Filename { get; set; }

        public IList<IScannedImage> Images { get; set; }

        private void exportPDFProcess()
        {
            var info = new PdfInfo
            {
                Title = "Scanned Image",
                Subject = "Scanned Image",
                Author = "NAPS2"
            };
            List<Image> imgs = Images.Select(x => (Image)x.GetImage()).ToList();

            try
            {
                pdfExporter.Export(Filename, imgs, info, num =>
                {
                    Invoke(new ThreadStart(() => SetStatus(num, imgs.Count)));
                    return true;
                });
            }
            catch (UnauthorizedAccessException)
            {
                errorOutput.DisplayError("You don't have permission to save files at this location.");
            }

            Invoke(new ThreadStart(Close));
        }

        void FPDFSave_Shown(object sender, EventArgs e)
        {
            new Thread(exportPDFProcess).Start();
        }

        public void SetStatus(int count, int total)
        {
            lblStatus.Text = string.Format(MiscResources.PdfStatus, count.ToString("G"), total.ToString("G"));
            Application.DoEvents();
        }
    }

}
