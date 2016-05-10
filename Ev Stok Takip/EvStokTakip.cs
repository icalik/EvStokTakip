using System;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.Windows.Forms;

namespace Ev_Stok_Takip
{
    public partial class EvStokTakip : Form
    {
        public OleDbConnection baglanti;
        public string secilenId;
        string hareketYon = "";
        public EvStokTakip()
        {
            InitializeComponent();
            baglanti = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\EvStokTakip_DB.accdb");
        }


        private void Form1_Load(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the 'evStokTakip_DBDataSet.Stok_Hareket' table. You can move, or remove it, as needed.
            this.stok_HareketTableAdapter.Fill(this.evStokTakip_DBDataSet.Stok_Hareket);
            // TODO: This line of code loads data into the 'evStokTakip_DBDataSet.Stok_Hareket' table. You can move, or remove it, as needed.
            this.stok_HareketTableAdapter.Fill(this.evStokTakip_DBDataSet.Stok_Hareket);

            // TODO: This line of code loads data into the 'evStokTakip_DBDataSet.Stok_Listesi' table. You can move, or remove it, as needed.
            this.stok_ListesiTableAdapter.Fill(this.evStokTakip_DBDataSet.Stok_Listesi);

        }

        private void btn_3Kaydet_Click(object sender, EventArgs e)
        {

            try
            {
                OleDbCommand komut = new OleDbCommand();
                komut.CommandType = CommandType.Text;
                komut.CommandText = @"insert into Stok_Listesi (Urun_Kodu,Urun_Adi,Birim,Bakiye,Aciklama) values ('" + urun_KoduTextBox.Text + "','" + urun_AdiTextBox.Text + "','" + birimComboBox.SelectedItem + "','" + bakiyeTextBox.Text + "','" + aciklamaTextBox.Text + "')";
                komut.Connection = baglanti;
                baglanti.Open();
                komut.ExecuteNonQuery();
                MessageBox.Show("OLDU!", "Oldu!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                baglanti.Close();
            }
            catch (Exception)
            {
                MessageBox.Show("Boş Geçemezsiniz!","Hata",MessageBoxButtons.OK,MessageBoxIcon.Error);

            }
            finally { }


        }





        public void secildi(object sender, DataGridViewCellMouseEventArgs e)
        {
            urun_KoduTextBox.Text = stok_ListesiDataGridView.SelectedRows[0].Cells[1].Value.ToString();
            urun_AdiTextBox.Text = stok_ListesiDataGridView.SelectedRows[0].Cells[2].Value.ToString();
            birimComboBox.SelectedIndex = birimComboBox.FindStringExact(stok_ListesiDataGridView.SelectedRows[0].Cells[3].Value.ToString());
            bakiyeTextBox.Text = stok_ListesiDataGridView.SelectedRows[0].Cells[4].Value.ToString();
            aciklamaTextBox.Text = stok_ListesiDataGridView.SelectedRows[0].Cells[5].Value.ToString();
            secilenId = stok_ListesiDataGridView.SelectedRows[0].Cells[0].Value.ToString();
            label8.Text = secilenId;
            ///
            txt_2UrunKodu.Text = stok_ListesiDataGridView.SelectedRows[0].Cells[1].Value.ToString();
            txt_2UrunAdi.Text = stok_ListesiDataGridView.SelectedRows[0].Cells[2].Value.ToString();
            cmb2_birim.SelectedIndex = cmb2_birim.FindStringExact(stok_ListesiDataGridView.SelectedRows[0].Cells[3].Value.ToString());
            //txt_2UrunMiktar.Text = stok_ListesiDataGridView.SelectedRows[0].Cells[4].Value.ToString();
            txt_2Aciklama.Text = stok_ListesiDataGridView.SelectedRows[0].Cells[5].Value.ToString();
        }

        private void yenile(object sender, EventArgs e)
        {
            this.stok_ListesiTableAdapter.Fill(this.evStokTakip_DBDataSet.Stok_Listesi);


        }

        public void button7_Click(object sender, EventArgs e)
        {
            try
            {
                OleDbCommand komut = new OleDbCommand();

                komut.CommandType = CommandType.Text;
                komut.CommandText = "UPDATE Stok_Listesi SET [Urun_Kodu]=@Urun_Kodu,[Urun_Adi]=@Urun_Adi,[Birim]=@Birim,[Bakiye]=@Bakiye,[Aciklama]=@Aciklama WHERE [ID]=@id";
                komut.Parameters.AddWithValue("@Urun_Kodu", urun_KoduTextBox.Text);
                komut.Parameters.AddWithValue("@Urun_Adi", urun_AdiTextBox.Text);
                try
                {
                    komut.Parameters.AddWithValue("@Birim", birimComboBox.SelectedItem.ToString());
                }
                catch (Exception)
                {

                    MessageBox.Show("Hata Olustu!", "HATA!", MessageBoxButtons.OK, MessageBoxIcon.Error);

                }

                komut.Parameters.AddWithValue("@Bakiye", bakiyeTextBox.Text);
                komut.Parameters.AddWithValue("@Aciklama", aciklamaTextBox.Text);
                komut.Parameters.AddWithValue("@id", secilenId);
                label8.Text = secilenId;
                komut.Connection = baglanti;
                baglanti.Open();
                try
                {
                    komut.ExecuteNonQuery();
                }
                catch (Exception)
                {
                    MessageBox.Show("Hata Olustu!", "HATA!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    
                }
                
                MessageBox.Show("Ürün Bilgisi Değiştirildi!", "Başarılı!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                baglanti.Close();
            }
            catch (SqlException)
            {
                MessageBox.Show("Hata Olustu!");
            }






        }

        private void button6_Click(object sender, EventArgs e)
        {
            try
            {
                OleDbCommand komut = new OleDbCommand();

                komut.CommandType = CommandType.Text;
                komut.CommandText = "DELETE FROM Stok_Listesi WHERE [ID]=@id";

                komut.Parameters.AddWithValue("@id", secilenId);
                label8.Text = secilenId;
                komut.Connection = baglanti;
                baglanti.Open();
                try
                {
                    komut.ExecuteNonQuery();
                }
                catch (Exception)
                {
                    MessageBox.Show("Hata Olustu!", "HATA!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    
                }
                
                MessageBox.Show("Ürün Kaydı Silinidi!", "Başarılı!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                baglanti.Close();
                urun_KoduTextBox.Clear();
                urun_AdiTextBox.Clear();
                birimComboBox.SelectedIndex = -1;
                bakiyeTextBox.Clear();
                aciklamaTextBox.Clear();
                label8.Text = "";
            }
            catch (SqlException)
            {
                MessageBox.Show("Hata Olustu!","HATA!",MessageBoxButtons.OK,MessageBoxIcon.Error);

            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            urun_KoduTextBox.Clear();
            urun_AdiTextBox.Clear();
            birimComboBox.SelectedIndex = -1;
            bakiyeTextBox.Clear();
            aciklamaTextBox.Clear();
            label8.Text = "";
        }

        private void groupBox3_Enter(object sender, EventArgs e)
        {

        }

        private void btn_2Kaydet_Click(object sender, EventArgs e)
        {

            bool hareketYonu = true;


            // rd_Giris.Checked = true;
            if (rd_Giris.Checked)
            {
                hareketYonu = true;
                hareketYon = "Giriş";
            }
            else if (rd_Cikis.Checked)
            {
                hareketYonu = false;
                hareketYon = "Çıkış";
            }

            try
            {
                OleDbCommand komut = new OleDbCommand();
                komut.CommandType = CommandType.Text;
                komut.CommandText = @"insert into Stok_Hareket (Urun_Kodu,Urun_Adi,Birim,Urun_Miktar,Aciklama,Yetkili, Islem_Tarihi,Hareket_Yonu) values ('" + txt_2UrunKodu.Text + "','" + txt_2UrunAdi.Text + "','" + cmb2_birim.SelectedItem + "','" + txt_2UrunMiktar.Text + "','" + txt_2Aciklama.Text + "','" + txt_2Yetkili.Text + "','" + dtp_2Tarih.Text + "','" + hareketYon + "')";
                komut.Connection = baglanti;
                baglanti.Open();
                komut.ExecuteNonQuery();
                baglanti.Close();
                ////



                komut.CommandText = "UPDATE Stok_Listesi SET[Bakiye]=@Bakiye WHERE [ID]=@id";
                double hareketBakiye;
                if (hareketYonu)
                {
                    hareketBakiye = Convert.ToDouble(bakiyeTextBox.Text) + Convert.ToDouble(txt_2UrunMiktar.Text);
                }
                else if (hareketYonu == false)
                {
                    hareketBakiye = Convert.ToDouble(bakiyeTextBox.Text) - Convert.ToDouble(txt_2UrunMiktar.Text);
                }
                else
                {
                    hareketBakiye = 0;
                }

                komut.Parameters.AddWithValue("@Bakiye", hareketBakiye);
                komut.Parameters.AddWithValue("@id", secilenId);
                label8.Text = secilenId;
                komut.Connection = baglanti;
                baglanti.Open();
                komut.ExecuteNonQuery();

                MessageBox.Show("Hareket Listesine Eklendi!", "Başarılı!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                baglanti.Close();

            }
            catch (Exception)
            {
                MessageBox.Show("Bos gecemezsiniz!","Hata!",MessageBoxButtons.OK,MessageBoxIcon.Error);

            }
            finally
            {

                this.stok_HareketTableAdapter.Fill(this.evStokTakip_DBDataSet.Stok_Hareket);

            }



        }

        private void çıkışToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void hakkındaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Hakkinda hakkinda = new Hakkinda();
            hakkinda.Show();
        }

        private void btn_2Iptal_Click(object sender, EventArgs e)
        {
            txt_2UrunAdi.Clear();
            txt_2UrunKodu.Clear();
            txt_2UrunMiktar.Clear();
            txt_2Aciklama.Clear();
            txt_2Yetkili.Clear();
            cmb2_birim.SelectedIndex = -1;
            rd_Cikis.Checked = false;
            rd_Giris.Checked = false;
        }

        private void stokListesiniExceleAktarToolStripMenuItem_Click(object sender, EventArgs e)
        {

            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            excel.Visible = true;
            Microsoft.Office.Interop.Excel.Workbook workbook = excel.Workbooks.Add(System.Reflection.Missing.Value);
            Microsoft.Office.Interop.Excel.Worksheet sheet1 = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[1];
            int StartCol = 1;
            int StartRow = 1;
            for (int j = 0; j < stok_ListesiDataGridView.Columns.Count; j++)
            {
                Microsoft.Office.Interop.Excel.Range myRange = (Microsoft.Office.Interop.Excel.Range)sheet1.Cells[StartRow, StartCol + j];
                myRange.Value2 = stok_ListesiDataGridView.Columns[j].HeaderText;
            }
            StartRow++;
            for (int i = 0; i < stok_ListesiDataGridView.Rows.Count; i++)
            {
                for (int j = 0; j < stok_ListesiDataGridView.Columns.Count; j++)
                {
                    try
                    {
                        Microsoft.Office.Interop.Excel.Range myRange = (Microsoft.Office.Interop.Excel.Range)sheet1.Cells[StartRow + i, StartCol + j];
                        myRange.Value2 = stok_ListesiDataGridView[j, i].Value == null ? "" : stok_ListesiDataGridView[j, i].Value;
                    }
                    catch { }
                }
            }
        }

        private void stokHareketListesiniExceleAktarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            excel.Visible = true;
            Microsoft.Office.Interop.Excel.Workbook workbook = excel.Workbooks.Add(System.Reflection.Missing.Value);
            Microsoft.Office.Interop.Excel.Worksheet sheet1 = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[1];
            int StartCol = 1;
            int StartRow = 1;
            for (int j = 0; j < stok_hareketDataGridView.Columns.Count; j++)
            {
                Microsoft.Office.Interop.Excel.Range myRange = (Microsoft.Office.Interop.Excel.Range)sheet1.Cells[StartRow, StartCol + j];
                myRange.Value2 = stok_hareketDataGridView.Columns[j].HeaderText;
            }
            StartRow++;
            for (int i = 0; i < stok_hareketDataGridView.Rows.Count; i++)
            {
                for (int j = 0; j < stok_hareketDataGridView.Columns.Count; j++)
                {
                    try
                    {
                        Microsoft.Office.Interop.Excel.Range myRange = (Microsoft.Office.Interop.Excel.Range)sheet1.Cells[StartRow + i, StartCol + j];
                        myRange.Value2 = stok_hareketDataGridView[j, i].Value == null ? "" : stok_hareketDataGridView[j, i].Value;
                    }
                    catch { }
                }
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {

            if (rd_1urunAdi.Checked)
            {
                string searchValue = txt_1Ara.Text;
                stok_ListesiDataGridView.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                try
                {
                    foreach (DataGridViewRow row in stok_ListesiDataGridView.Rows)
                    {
                        if (row.Cells[2].Value.ToString().Equals(searchValue))
                        {
                            row.Selected = true;
                            break;
                        }
                    }
                }
                catch (Exception exc)
                {
                    MessageBox.Show(exc.Message);
                }
            }
            else if (rd_1urunKodu.Checked)
            {
                string searchValue = txt_1Ara.Text;
                stok_ListesiDataGridView.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                try
                {
                    foreach (DataGridViewRow row in stok_ListesiDataGridView.Rows)
                    {
                        if (row.Cells[1].Value.ToString().Equals(searchValue))
                        {
                            row.Selected = true;
                            break;
                        }
                    }
                }
                catch (Exception exc)
                {
                    MessageBox.Show(exc.Message);
                }
            }
        }
    }
}







