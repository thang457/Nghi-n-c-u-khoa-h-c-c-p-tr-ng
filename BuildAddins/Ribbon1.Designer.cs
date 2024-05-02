namespace BuildAddins
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.button1 = this.Factory.CreateRibbonButton();
            this.button2 = this.Factory.CreateRibbonButton();
            this.button3 = this.Factory.CreateRibbonButton();
            this.button4 = this.Factory.CreateRibbonButton();
            this.button5 = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.button6 = this.Factory.CreateRibbonButton();
            this.button7 = this.Factory.CreateRibbonButton();
            this.button8 = this.Factory.CreateRibbonButton();
            this.button9 = this.Factory.CreateRibbonButton();
            this.button10 = this.Factory.CreateRibbonButton();
            this.button11 = this.Factory.CreateRibbonButton();
            this.button12 = this.Factory.CreateRibbonButton();
            this.button13 = this.Factory.CreateRibbonButton();
            this.button14 = this.Factory.CreateRibbonButton();
            this.button15 = this.Factory.CreateRibbonButton();
            this.button16 = this.Factory.CreateRibbonButton();
            this.button17 = this.Factory.CreateRibbonButton();
            this.button18 = this.Factory.CreateRibbonButton();
            this.button19 = this.Factory.CreateRibbonButton();
            this.button20 = this.Factory.CreateRibbonButton();
            this.button21 = this.Factory.CreateRibbonButton();
            this.button22 = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.btnInsertImage = this.Factory.CreateRibbonButton();
            this.btnTenBang = this.Factory.CreateRibbonButton();
            this.btnTenAnh = this.Factory.CreateRibbonButton();
            this.btnCongThucToan = this.Factory.CreateRibbonButton();
            this.btnGhiChu = this.Factory.CreateRibbonButton();
            this.btnprogcode = this.Factory.CreateRibbonButton();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.button23 = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.group4.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Groups.Add(this.group4);
            this.tab1.Label = "Add- Inss";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.button1);
            this.group1.Items.Add(this.button2);
            this.group1.Items.Add(this.button3);
            this.group1.Items.Add(this.button4);
            this.group1.Items.Add(this.button5);
            this.group1.Label = "Định dạng văn bản";
            this.group1.Name = "group1";
            // 
            // button1
            // 
            this.button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button1.Image = global::BuildAddins.Properties.Resources.Screenshot_2024_02_23_145319;
            this.button1.Label = "Tiêu đề";
            this.button1.Name = "button1";
            this.button1.ShowImage = true;
            // 
            // button2
            // 
            this.button2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button2.Image = global::BuildAddins.Properties.Resources.Screenshot_2024_02_23_145618;
            this.button2.Label = "Tác giả";
            this.button2.Name = "button2";
            this.button2.ShowImage = true;
            // 
            // button3
            // 
            this.button3.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button3.Image = global::BuildAddins.Properties.Resources.Screenshot_2024_02_23_145717;
            this.button3.Label = "Địa chỉ";
            this.button3.Name = "button3";
            this.button3.ShowImage = true;
            // 
            // button4
            // 
            this.button4.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button4.Image = global::BuildAddins.Properties.Resources.Screenshot_2024_02_23_150321;
            this.button4.Label = "Tóm tắt";
            this.button4.Name = "button4";
            this.button4.ShowImage = true;
            // 
            // button5
            // 
            this.button5.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button5.Image = global::BuildAddins.Properties.Resources.Screenshot_2024_02_23_145733;
            this.button5.Label = "Từ khóa";
            this.button5.Name = "button5";
            this.button5.ShowImage = true;
            // 
            // group2
            // 
            this.group2.Items.Add(this.button6);
            this.group2.Items.Add(this.button7);
            this.group2.Items.Add(this.button8);
            this.group2.Items.Add(this.button9);
            this.group2.Items.Add(this.button10);
            this.group2.Items.Add(this.button11);
            this.group2.Items.Add(this.button12);
            this.group2.Items.Add(this.button13);
            this.group2.Items.Add(this.button14);
            this.group2.Items.Add(this.button15);
            this.group2.Items.Add(this.button16);
            this.group2.Items.Add(this.button17);
            this.group2.Items.Add(this.button18);
            this.group2.Items.Add(this.button19);
            this.group2.Items.Add(this.button20);
            this.group2.Items.Add(this.button21);
            this.group2.Items.Add(this.button22);
            this.group2.Label = "Danh mục";
            this.group2.Name = "group2";
            // 
            // button6
            // 
            this.button6.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button6.Image = global::BuildAddins.Properties.Resources.Screenshot_2024_02_23_145746;
            this.button6.Label = "Tiêu đề cấp 1";
            this.button6.Name = "button6";
            this.button6.ShowImage = true;
            // 
            // button7
            // 
            this.button7.Label = "Tiêu đề cấp 2";
            this.button7.Name = "button7";
            // 
            // button8
            // 
            this.button8.Label = "Tiêu đề cấp 3";
            this.button8.Name = "button8";
            // 
            // button9
            // 
            this.button9.Label = "Tiêu đề cấp 4";
            this.button9.Name = "button9";
            // 
            // button10
            // 
            this.button10.Image = global::BuildAddins.Properties.Resources.Screenshot_2024_02_23_151951;
            this.button10.Label = "Chấm";
            this.button10.Name = "button10";
            this.button10.ShowImage = true;
            // 
            // button11
            // 
            this.button11.Image = global::BuildAddins.Properties.Resources.Screenshot_2024_02_23_151854;
            this.button11.Label = "Trừ";
            this.button11.Name = "button11";
            this.button11.ShowImage = true;
            // 
            // button12
            // 
            this.button12.Image = global::BuildAddins.Properties.Resources.Screenshot_2024_02_23_150804;
            this.button12.Label = "Số";
            this.button12.Name = "button12";
            this.button12.ShowImage = true;
            // 
            // button13
            // 
            this.button13.Image = global::BuildAddins.Properties.Resources.tải_xuống__1____Copy;
            this.button13.Label = "Tăng level +";
            this.button13.Name = "button13";
            this.button13.ShowImage = true;
            // 
            // button14
            // 
            this.button14.Image = global::BuildAddins.Properties.Resources.tải_xuống__1_;
            this.button14.Label = "Giảm level -";
            this.button14.Name = "button14";
            this.button14.ShowImage = true;
            // 
            // button15
            // 
            this.button15.Label = "1../..n..";
            this.button15.Name = "button15";
            // 
            // button16
            // 
            this.button16.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button16.Image = global::BuildAddins.Properties.Resources.Screenshot_2024_02_23_145808;
            this.button16.Label = "Chữ bình thường";
            this.button16.Name = "button16";
            this.button16.ShowImage = true;
            // 
            // button17
            // 
            this.button17.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button17.Image = global::BuildAddins.Properties.Resources.Screenshot_2024_02_23_1458181;
            this.button17.Label = "1 cột";
            this.button17.Name = "button17";
            this.button17.ShowImage = true;
            // 
            // button18
            // 
            this.button18.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button18.Image = global::BuildAddins.Properties.Resources.Screenshot_2024_02_23_145826;
            this.button18.Label = "2 cột";
            this.button18.Name = "button18";
            this.button18.ShowImage = true;
            // 
            // button19
            // 
            this.button19.Label = "Tạo khoảng cách";
            this.button19.Name = "button19";
            // 
            // button20
            // 
            this.button20.Label = "Xóa khoảng cách";
            this.button20.Name = "button20";
            // 
            // button21
            // 
            this.button21.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button21.Image = global::BuildAddins.Properties.Resources.Screenshot_2024_02_23_145857;
            this.button21.Label = "Tạo chú thích";
            this.button21.Name = "button21";
            this.button21.ShowImage = true;
            // 
            // button22
            // 
            this.button22.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button22.Image = global::BuildAddins.Properties.Resources.Screenshot_2024_02_23_145909;
            this.button22.Label = "Tài liệu tham khảo";
            this.button22.Name = "button22";
            this.button22.ShowImage = true;
            // 
            // group3
            // 
            this.group3.Items.Add(this.btnInsertImage);
            this.group3.Items.Add(this.btnTenBang);
            this.group3.Items.Add(this.btnTenAnh);
            this.group3.Items.Add(this.btnCongThucToan);
            this.group3.Items.Add(this.btnGhiChu);
            this.group3.Items.Add(this.btnprogcode);
            this.group3.Label = "Ảnh, bảng, công thức toán";
            this.group3.Name = "group3";
            // 
            // btnInsertImage
            // 
            this.btnInsertImage.Image = global::BuildAddins.Properties.Resources.Screenshot_2024_02_23_1511231;
            this.btnInsertImage.Label = "Chèn ảnh";
            this.btnInsertImage.Name = "btnInsertImage";
            this.btnInsertImage.ShowImage = true;
            this.btnInsertImage.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnInsertImage_Click);
            // 
            // btnTenBang
            // 
            this.btnTenBang.Image = global::BuildAddins.Properties.Resources.tải_xuống__3_;
            this.btnTenBang.Label = "Tên bảng";
            this.btnTenBang.Name = "btnTenBang";
            this.btnTenBang.ShowImage = true;
            this.btnTenBang.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnTenBang_Click);
            // 
            // btnTenAnh
            // 
            this.btnTenAnh.Image = global::BuildAddins.Properties.Resources.tải_xuống__2_;
            this.btnTenAnh.Label = "Tên ảnh";
            this.btnTenAnh.Name = "btnTenAnh";
            this.btnTenAnh.ShowImage = true;
            this.btnTenAnh.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnTenAnh_Click);
            // 
            // btnCongThucToan
            // 
            this.btnCongThucToan.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnCongThucToan.Image = global::BuildAddins.Properties.Resources.tải_xuống;
            this.btnCongThucToan.Label = "Công thức toán";
            this.btnCongThucToan.Name = "btnCongThucToan";
            this.btnCongThucToan.ShowImage = true;
            // 
            // btnGhiChu
            // 
            this.btnGhiChu.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnGhiChu.Image = global::BuildAddins.Properties.Resources.Screenshot_2024_02_23_145939;
            this.btnGhiChu.Label = "Ghi chú";
            this.btnGhiChu.Name = "btnGhiChu";
            this.btnGhiChu.ShowImage = true;
            // 
            // btnprogcode
            // 
            this.btnprogcode.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnprogcode.Image = global::BuildAddins.Properties.Resources.Screenshot_2024_02_23_145951;
            this.btnprogcode.Label = "Progcode";
            this.btnprogcode.Name = "btnprogcode";
            this.btnprogcode.ShowImage = true;
            // 
            // group4
            // 
            this.group4.Items.Add(this.button23);
            this.group4.Label = "group4";
            this.group4.Name = "group4";
            // 
            // button23
            // 
            this.button23.Label = "check ảnh";
            this.button23.Name = "button23";
            this.button23.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button23_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.group4.ResumeLayout(false);
            this.group4.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button5;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button6;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button7;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button8;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button9;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button10;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button11;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button12;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button13;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button14;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button15;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button16;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button17;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button18;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button19;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button20;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button21;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button22;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnInsertImage;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnTenBang;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnTenAnh;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCongThucToan;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGhiChu;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnprogcode;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button23;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
