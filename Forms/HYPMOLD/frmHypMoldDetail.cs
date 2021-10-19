using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BMS
{

	public partial class frmHypMoldDetail : Form
	{
		public string _MaHang = "";
		public int _KnifeDetailListID;
		public int _ID;
		public bool _isNew;
		int _selectedNumber;
		public int _ischeck;
		public int KnifeID = -1;
		public string Createby;
		public string Updateby;
		public string CongDoan;
		object X;
		public bool _isNewTextBox;

		public frmHypMoldDetail()
		{
			InitializeComponent();
		}

		private void frmHypMoldDetail_Load(object sender, EventArgs e)
		{
			txtMahang.Text = _MaHang;
			txtMin.Enabled = txtMax.Enabled = txtStandardValue.Enabled = txtX.Enabled = txtY.Enabled = false;
			if (!_isNew)
			{
				loadHypMoldDetail();
				lbl.Text = "Người cập nhật";
				txtCreateby.Text = Updateby;
			}
			else
			{
				lbl.Text = "Người tạo";
				txtCreateby.Text = Createby;
			}
		}

		private void loadHypMoldDetail()
		{
			string sql = $"Select * from dbo.KnifeJigWorking where ID = {_ID}";
			DataTable dt = TextUtilsHP.Select(sql);
			txtCheckingName.Text = TextUtilsHP.ToString(dt.Rows[0]["WorkingName"]);
			txtStt.Text = TextUtilsHP.ToString(dt.Rows[0]["SortOrder"]);
			cboType.SelectedIndex = TextUtilsHP.ToInt(dt.Rows[0]["CheckValueType"]);
			txtStandardValue.Text = TextUtilsHP.ToString(dt.Rows[0]["PeriodValue"]);
			txtMin.Text = TextUtilsHP.ToString(dt.Rows[0]["MinValue"]);
			txtMax.Text = TextUtilsHP.ToString(dt.Rows[0]["MaxValue"]);
			txtUnit.Text = TextUtilsHP.ToString(dt.Rows[0]["Unit"]);
			txtX.Text = TextUtilsHP.ToString(dt.Rows[0]["X"]);
			txtY.Text = TextUtilsHP.ToString(dt.Rows[0]["Y"]);
			checkBox1.Checked = TextUtilsHP.ToBoolean(dt.Rows[0]["isDisplay"]);
			X = dt.Rows[0]["X"];
		}

		private void cboType_SelectedIndexChanged(object sender, EventArgs e)
		{
			_selectedNumber = cboType.SelectedIndex;
			if (_selectedNumber == 1)
			{
				txtStandardValue.Enabled = false;
				txtMin.Enabled = txtMax.Enabled = true;
				txtUnit.Enabled = true;
				txtMin.Focus();
			}
			else if (_selectedNumber == 2)
			{
				txtMin.Enabled = txtMax.Enabled = false;
				txtStandardValue.Enabled = true;
				txtUnit.Enabled = false;
				txtStandardValue.Focus();
			}
			else
			{
				txtMin.Enabled = txtMax.Enabled = false;
				txtStandardValue.Enabled = false;
				txtUnit.Enabled = false;
			}
		}

		private void btnSaveClose_Click(object sender, EventArgs e)
		{
			if (_selectedNumber == 1)
			{
				if (txtCheckingName.Text == "" || txtStt.Text == "" || txtUnit.Text == "" || txtMin.Text == "" || txtMax.Text == "")
				{
					MessageBox.Show("Vui lòng điền đầy đủ thông tin!", "WARNING", MessageBoxButtons.OK, MessageBoxIcon.Warning);
					return;
				}
			}
			else if (_selectedNumber == 2)
			{
				if (txtCheckingName.Text == "" || txtStt.Text == "" || txtStandardValue.Text == "")
				{
					MessageBox.Show("Vui lòng điền đầy đủ thông tin!", "WARNING", MessageBoxButtons.OK, MessageBoxIcon.Warning);
					return;
				}
			}
			else if (_selectedNumber == 0)
			{
				MessageBox.Show("Vui lòng điền đầy đủ thông tin!", "WARNING", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				return;
			}
			string sqlSTT = $"Select SortOrder from dbo.KnifeJigWorking where STT ={TextUtilsHP.ToInt(txtStt.Text.Trim())} and ID <> {_ID} and KnifeDetailListID = {_KnifeDetailListID}";
			DataTable dt = TextUtilsHP.Select(sqlSTT);

			if (dt.Rows.Count > 0)
			{
				MessageBox.Show("STT này đã tồn tại!", "WARNING", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				return;
			}
			string sql = "";
			string condition;
			if (_selectedNumber == 2)
			{
				condition = txtStandardValue.Text.Trim();
			}
			else
			{
				condition = txtStandardValue.Text = txtMin.Text.Trim() + "~" + txtMax.Text.Trim();

			}
			if (_ischeck == 1 && X == null)
			{
				_isNewTextBox = true;
			}
			else
			{
				_isNewTextBox = false;
			}
			if (_isNew)
			{
				sql = $"insert into dbo.KnifeJigWorking(KnifeDetailListID,WorkingName,ValueType,PeriodValue,MinValue,MaxValue,Unit,SortOrder,isDisplay,Condition,CreatedBy,CreatedDate,CheckValueType,X,Y,ProductStepCode)" +
				$" Values({_KnifeDetailListID}," +
				$"N'{txtCheckingName.Text.Trim()}'," +
				$"{1}," +
				$"'{txtStandardValue.Text.Trim()}'," +
				$"{TextUtilsHP.ToDecimal(txtMin.Text.Trim())}," +
				$"{TextUtilsHP.ToDecimal(txtMax.Text.Trim())}," +
				$"'{txtUnit.Text.Trim()}'," +
				$"{TextUtilsHP.ToInt(txtStt.Text.Trim())}," +
				$"{_ischeck}," +
				$" '{condition}'," +
				$" '{txtCreateby.Text.Trim()}'," +
				$" '{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")}'," +
				$"{_selectedNumber}," +
				$"{TextUtilsHP.ToInt(txtX.Text.Trim())}," +
				$"{TextUtilsHP.ToInt(txtY.Text.Trim())}," +
				$" '{CongDoan}')" +
				$"select SCOPE_IDENTITY() as IDD";
			}
			else
			{
				sql = $"update dbo.KnifeJigWorking set  " +
					// $" KnifeDetailListID = '{_KnifeDetailListID}'" +
					$" WorkingName = N'{txtCheckingName.Text.Trim()}'" +
					$", CheckValueType ={_selectedNumber}" +
					$", PeriodValue = '{txtStandardValue.Text.Trim()}'" +
					$", MinValue = {TextUtilsHP.ToDecimal(txtMin.Text.Trim())}" +
					$", MaxValue = {TextUtilsHP.ToDecimal(txtMax.Text.Trim())}" +
					$", Unit= '{txtUnit.Text.Trim()}'" +
					$", SortOrder = {TextUtilsHP.ToInt(txtStt.Text.Trim())}" +
					$", isDisplay = {_ischeck}" +
					$", ProductStepCode = '{CongDoan}' " +
					$", X = {TextUtilsHP.ToInt(txtX.Text.Trim())}" +
					$", Y = {TextUtilsHP.ToInt(txtY.Text.Trim())}" +
					$", UpdatedBy = '{txtCreateby.Text.Trim()}' " +
					$", Condition = '{condition}' " +
					$", UpdatedDate = '{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")}'" +
					$" Where ID = {_ID}";
			}
			//KnifeID= TextUtilsHP.ToInt(TextUtilsHP.ExcuteSQL(sql));
			KnifeID = TextUtilsHP.ToInt(TextUtilsHP.ExcuteScalar(sql));
			this.DialogResult = DialogResult.OK;
		}

		private void checkBox1_CheckedChanged(object sender, EventArgs e)
		{
			if (checkBox1.Checked)
			{
				_ischeck = 1;
				txtY.Enabled = true;
				txtX.Enabled = true;
			}
			else
			{
				_ischeck = 0;
				txtY.Enabled = false;
				txtX.Enabled = false;
			}
		}

		private void frmHypMoldDetail_FormClosed(object sender, FormClosedEventArgs e)
		{
			this.DialogResult = DialogResult.OK;
		}

		private void txtCheckingName_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode != Keys.Enter) return;
			txtStt.Focus();
		}

		private void txtStt_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode != Keys.Enter) return;
			cboType.Focus();
		}

		private void txtStandardValue_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				txtX.Focus();
			}
		}

		private void txtMin_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode != Keys.Enter) return;
			txtMax.Focus();
		}

		private void txtMax_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode != Keys.Enter) return;
			txtUnit.Focus();
		}

		private void cấtToolStripMenuItem_Click(object sender, EventArgs e)
		{
			btnSaveClose_Click(null, null);
		}

		private void txtMin_KeyPress(object sender, KeyPressEventArgs e)
		{
			//if (!Char.IsDigit(e.KeyChar) && !Char.IsControl(e.KeyChar))
			//	e.Handled = true;
		}

		private void txtMax_KeyPress(object sender, KeyPressEventArgs e)
		{
			//if (!Char.IsDigit(e.KeyChar) && !Char.IsControl(e.KeyChar))
			//	e.Handled = true;
		}

		private void txtUnit_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				txtX.Focus();
			}
		}

		private void txtX_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				txtY.Focus();
			}
		}
	}
}
