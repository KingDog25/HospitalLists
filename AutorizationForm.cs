using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace HospitalLists
{
    public partial class AutorizationForm1 : Form
    {
        public AutorizationForm1()
        {
            InitializeComponent();
        }

        private void buttonAutOK_Click(object sender, EventArgs e)
        {
            string login = textBoxAutLogin.Text;
            string password = textBoxAutPass.Text;

            if (login == "test" && password == "test" && comboBoxAut.SelectedIndex!=-1)
            {
                // Установка результата авторизации и закрытие формы
                this.DialogResult = DialogResult.OK;
                this.Close();
                SingletonClass autorizObject = SingletonClass.getInstance();
                if (comboBoxAut.SelectedIndex ==0)
                    autorizObject.setField1(2);     //2-тип авторизации для врача
                if (comboBoxAut.SelectedIndex == 1)
                    autorizObject.setField1(1);     //1-тип авторизации для руководства
                if (comboBoxAut.SelectedIndex == 2)
                    autorizObject.setField1(3);     //3-тип авторизации для медсестры
                this.Close();
            }
            else
            {
                MessageBox.Show("Неверные логин или пароль");
            }
        }
    }
}
