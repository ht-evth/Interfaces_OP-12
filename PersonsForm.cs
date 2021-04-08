using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Catering_OP_6 {
	public partial class PersonsForm : Form {

		//Form Owner;

		List<PlaceHolderTextBox> placeHolderTextBoxes;


		public PersonsForm(Form form) {
			InitializeComponent();

			//Owner = form;

			placeHolderTextBoxes = new List<PlaceHolderTextBox>() {
				TextBox_Post_FinPerson,

				TextBox_Member1_FullName,
				TextBox_Member2_FullName,
				TextBox_Member3_FullName,

				TexBox_Member2_Post,
				TexBox_Member3_Post,

				TextBox_FullName_CheckPerson,
				TextBox_Cashier_FullName
			};
		}

		public void ClearForm() {
			// текстбоксы с плейсхолдерами - пустая строка и установить плейсхолдер
			foreach (var item in placeHolderTextBoxes) {
				item.Text = "";
				item.setPlaceholder();
			}
		}

		public List<string> GetPosts() {
			return new List<string>() {
				TextBox_Post_FinPerson.Text,
				textBox_Member1_Post.Text,
				TexBox_Member2_Post.Text,
				TexBox_Member3_Post.Text,
			};
		}

		public List<string> GetFullNames() {
			return new List<string>() {
				TextBox_Member1_FullName.Text,
				TextBox_Member2_FullName.Text,
				TextBox_Member3_FullName.Text,
				TextBox_Cashier_FullName.Text,
				TextBox_FullName_CheckPerson.Text
			};
		}

		private void button_Save_Click(object sender, EventArgs e) {
			this.Hide();
		}

        private void PersonsForm_Load(object sender, EventArgs e)
        {

        }

        private void PersonsForm_FormClosing(object sender, FormClosingEventArgs e)
        {
			e.Cancel = true;
			this.Hide();
        }

        private void button1_Click(object sender, EventArgs e)
        {
			this.Hide();
        }
    }
}
