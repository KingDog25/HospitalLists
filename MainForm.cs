using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;
using System.Configuration;
using System.Media;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Net.Mail;


namespace HospitalLists
{
    public partial class MainForm : Form
    {
        // Создаем экземпляр класса SoundPlayer и указываем путь к аудиофайлу
        SoundPlayer sound1 = new SoundPlayer(@"..\..\sound\Запрос.wav");
        SoundPlayer soundExit = new SoundPlayer(@"..\..\sound\Exit.wav");
        SoundPlayer soundAdd = new SoundPlayer(@"..\..\sound\lineAdd.wav");
        SoundPlayer soundDrum = new SoundPlayer(@"..\..\sound\DRUMROLL.WAV");
        SoundPlayer soundPush = new SoundPlayer(@"..\..\sound\PUSH.WAV");
        SoundPlayer soundWHOOSH = new SoundPlayer(@"..\..\sound\WHOOSH.WAV");

        public MainForm()
        {
            InitializeComponent();

        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            // Создание экземпляра формы AuthorizationForm1
            AutorizationForm1 authorizationForm = new AutorizationForm1();
            authorizationForm.ShowDialog();
            больничныйЛистToolStripMenuItem.Enabled = false;
            больничныйЛистToolStripMenuItem.Visible = false;
            // Проверка результата авторизации
            if (authorizationForm.DialogResult == DialogResult.OK)
            {
                // Отображение главной формы
                this.Show();
            }
            else
            {
                // Закрытие приложения, если авторизация не прошла успешно
                Application.Exit();
            }
            tabPagesRemoveAdd(0);

            SingletonClass autorizObject = SingletonClass.getInstance();
            if (autorizObject.getField1() == 2) //если авторизован врач
            {
                учрежденияОтделыИДолжностиToolStripMenuItem.Visible = false;
                сотрудникиToolStripMenuItem.Visible = false;
            }

            if (autorizObject.getField1() == 3) //если авторизована медсестра
            {
                учрежденияОтделыИДолжностиToolStripMenuItem.Visible = false;
                сотрудникиToolStripMenuItem.Visible = false;
                пациентыИОпекуныToolStripMenuItem.Visible = false;
                больничныеЛистыToolStripMenuItem.Visible = false;
                спецКодыToolStripMenuItem.Visible = false;
                buttonDelete.Visible = false;
                //buttonUpdate.Visible = false;
                buttonAdd.Visible = false;
            }
        }

        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            soundExit.Play();
            Application.Exit();
        }

        void tabPagesRemoveAdd(int addRow)
        {
            tabControl1.TabPages.Remove(tabPage1);
            tabControl1.TabPages.Remove(tabPage2);
            tabControl1.TabPages.Remove(tabPage3);
            tabControl1.TabPages.Remove(tabPage4);
            tabControl1.TabPages.Remove(tabPage5);

            if (addRow == 1)
                tabControl1.TabPages.Add(tabPage1);

            if (addRow == 2)
            {
                tabControl1.TabPages.Add(tabPage1);
                tabControl1.TabPages.Add(tabPage2);
            }

            if (addRow == 3)
            {
                tabControl1.TabPages.Add(tabPage1);
                tabControl1.TabPages.Add(tabPage2);
                tabControl1.TabPages.Add(tabPage3);
            }

            if (addRow == 4)
            {
                tabControl1.TabPages.Add(tabPage1);
                tabControl1.TabPages.Add(tabPage2);
                tabControl1.TabPages.Add(tabPage3);
                tabControl1.TabPages.Add(tabPage4);
            }

            if (addRow == 5)
            {
                tabControl1.TabPages.Add(tabPage1);
                tabControl1.TabPages.Add(tabPage2);
                tabControl1.TabPages.Add(tabPage3);
                tabControl1.TabPages.Add(tabPage4);
                tabControl1.TabPages.Add(tabPage5);
            }
        }

        void resizecolumns()
        {
            dataGridView1.AutoResizeColumns();
            dataGridView2.AutoResizeColumns();
            dataGridView3.AutoResizeColumns();
            dataGridView4.AutoResizeColumns();
            dataGridView5.AutoResizeColumns();
        }

        private void учрежденияОтделыИДолжностиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            tabPagesRemoveAdd(3);
            SingletonClass autorizObject = SingletonClass.getInstance();
            if (autorizObject.getField1() == 1) //если авторизован руководитель
            {
                //Учреждение
                using (SqlConnection sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["HospitalLists"].ConnectionString))
                {
                    sqlConnection.Open();
                    SqlDataAdapter dataAdapter = new SqlDataAdapter("SELECT ID AS 'Код учреждения', MedicalName AS 'Название', Adress AS 'Адрес',  ORGN AS 'Код ОРГН' From Medical_Institution", sqlConnection);
                    DataSet dataSet = new DataSet();     //создаем датасет для результатов выборки
                    //Запишем данные в таблицу формы
                    dataAdapter.Fill(dataSet);          //заполняем датасет с помощью адаптера
                    dataGridView1.DataSource = dataSet.Tables[0];
                    Table.tableName = "Table_Medical_Institution";
                    tabPage1.Text = "Учреждения";
                }

                //Отделение
                using (SqlConnection sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["HospitalLists"].ConnectionString))
                {
                    sqlConnection.Open();
                    SqlDataAdapter dataAdapter = new SqlDataAdapter("SELECT d.ID AS 'Код отделения', d.Department AS 'Название отделения', m.MedicalName AS 'Название учреждения'\r\nFROM Department d\r\nJOIN Medical_Institution m ON d.MedicalInst_ID = m.ID;", sqlConnection);
                    DataSet dataSet = new DataSet();     //создаем датасет для результатов выборки
                    //Запишем данные в таблицу формы
                    dataAdapter.Fill(dataSet);          //заполняем датасет с помощью адаптера
                    dataGridView2.DataSource = dataSet.Tables[0];
                    Table.tableName = "Table_Department";
                    tabPage2.Text = "Отделения";
                }

                //Должность
                using (SqlConnection sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["HospitalLists"].ConnectionString))
                {
                    sqlConnection.Open();
                    SqlDataAdapter dataAdapter = new SqlDataAdapter("SELECT p.ID AS 'Код должности', p.PostName AS 'Название должности', d.Department AS 'Отдел'\r\nFROM Post p\r\n JOIN Department d ON p.Department_ID = d.ID;", sqlConnection);
                    DataSet dataSet = new DataSet();     //создаем датасет для результатов выборки
                    //Запишем данные в таблицу формы
                    dataAdapter.Fill(dataSet);          //заполняем датасет с помощью адаптера
                    dataGridView3.DataSource = dataSet.Tables[0];
                    Table.tableName = "Table_Post";
                    tabPage3.Text = "Должности";
                }
                resizecolumns(); //автоподбор ширины столбцов

                // Воспроизводим звук
                sound1.Play();
            }
        }

        private void спецКодыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            tabPagesRemoveAdd(5);
            SingletonClass autorizObject = SingletonClass.getInstance();
            if (autorizObject.getField1() == 1 || autorizObject.getField1() == 2) //если авторизован руководитель или врач
            {
                //Код причины 1
                using (SqlConnection sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["HospitalLists"].ConnectionString))
                {
                    sqlConnection.Open();
                    SqlDataAdapter dataAdapter = new SqlDataAdapter("SELECT ID AS 'Код причины __', Description AS 'Описание причины' From Reason_1", sqlConnection);
                    DataSet dataSet = new DataSet();     //создаем датасет для результатов выборки
                    //Запишем данные в таблицу формы
                    dataAdapter.Fill(dataSet);          //заполняем датасет с помощью адаптера
                    dataGridView1.DataSource = dataSet.Tables[0];
                    Table.tableName = "Table_Reason_1";
                    tabPage1.Text = "Причина (основная)";
                }

                //Код причины 2
                using (SqlConnection sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["HospitalLists"].ConnectionString))
                {
                    sqlConnection.Open();
                    SqlDataAdapter dataAdapter = new SqlDataAdapter("SELECT Code AS 'Код причины ___', Description AS 'Описание причины' From Rereason_2", sqlConnection);
                    DataSet dataSet = new DataSet();     //создаем датасет для результатов выборки
                    //Запишем данные в таблицу формы
                    dataAdapter.Fill(dataSet);          //заполняем датасет с помощью адаптера
                    dataGridView2.DataSource = dataSet.Tables[0];
                    Table.tableName = "Table_Rereason_2";
                    tabPage2.Text = "Причина (дополнительная)";
                }

                //Нарушение режима
                using (SqlConnection sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["HospitalLists"].ConnectionString))
                {
                    sqlConnection.Open();
                    SqlDataAdapter dataAdapter = new SqlDataAdapter("SELECT Code AS 'Код отмены __', Description AS 'Описание нарушения режима' From DisturbanceRegime", sqlConnection);
                    DataSet dataSet = new DataSet();     //создаем датасет для результатов выборки
                    //Запишем данные в таблицу формы
                    dataAdapter.Fill(dataSet);          //заполняем датасет с помощью адаптера
                    dataGridView3.DataSource = dataSet.Tables[0];
                    Table.tableName = "Table_DisturbanceRegime";
                    tabPage3.Text = "Нарушение режима";
                }

                //Родственные связи
                using (SqlConnection sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["HospitalLists"].ConnectionString))
                {
                    sqlConnection.Open();
                    SqlDataAdapter dataAdapter = new SqlDataAdapter("SELECT Code AS 'Код связи __', Description AS 'Описание родственной связи' From Care_Code_Table", sqlConnection);
                    DataSet dataSet = new DataSet();     //создаем датасет для результатов выборки
                    //Запишем данные в таблицу формы
                    dataAdapter.Fill(dataSet);          //заполняем датасет с помощью адаптера
                    dataGridView4.DataSource = dataSet.Tables[0];
                    Table.tableName = "Table_Care_Code_Table";
                    tabPage4.Text = "Родственные связи";
                }

                //Иное
                using (SqlConnection sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["HospitalLists"].ConnectionString))
                {
                    sqlConnection.Open();
                    SqlDataAdapter dataAdapter = new SqlDataAdapter("SELECT Code AS 'Код Иное __', Description AS 'Описание' From Inoe_Codes", sqlConnection);
                    DataSet dataSet = new DataSet();     //создаем датасет для результатов выборки
                    //Запишем данные в таблицу формы
                    dataAdapter.Fill(dataSet);          //заполняем датасет с помощью адаптера
                    dataGridView5.DataSource = dataSet.Tables[0];
                    Table.tableName = "Table_Inoe_Codes";
                    tabPage5.Text = "Иное";
                }
                resizecolumns(); //автоподбор ширины столбцов
                // Воспроизводим звук
                sound1.Play();
            }
        }


        private void сотрудникиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            tabPagesRemoveAdd(1);
            tabPage1.Text = "Врачи";
            SingletonClass autorizObject = SingletonClass.getInstance();
            if (autorizObject.getField1() == 1 || autorizObject.getField1() == 2) //если авторизован руководитель или врач
            {
                //Врачи
                using (SqlConnection sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["HospitalLists"].ConnectionString))
                {
                    sqlConnection.Open();
                    SqlDataAdapter dataAdapter = new SqlDataAdapter("SELECT d.ID AS 'Код сотрудника', LastName AS 'Фамилия', Name AS 'Имя', Patronymic AS 'Отчество', Phone AS 'Телефон', PostName AS 'Должность' From Doctor d JOIN Post p ON p.ID = d.Post_ID;", sqlConnection);
                    DataSet dataSet = new DataSet();     //создаем датасет для результатов выборки
                    //Запишем данные в таблицу формы
                    dataAdapter.Fill(dataSet);          //заполняем датасет с помощью адаптера
                    dataGridView1.DataSource = dataSet.Tables[0];
                    Table.tableName = "Table_Doctor";
                    tabPage1.Text = "Врачи";
                }
            }
            resizecolumns(); //автоподбор ширины столбцов
            // Воспроизводим звук
            sound1.Play();
        }

        private void пациентыИОпекуныToolStripMenuItem_Click(object sender, EventArgs e)
        {
            tabPagesRemoveAdd(2);
            SingletonClass autorizObject = SingletonClass.getInstance();
            if (autorizObject.getField1() == 1 || autorizObject.getField1() == 2) //если авторизован руководитель или врач
            {
                //Пациенты
                using (SqlConnection sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["HospitalLists"].ConnectionString))
                {
                    sqlConnection.Open();
                    SqlDataAdapter dataAdapter = new SqlDataAdapter("SELECT ID AS 'ИД Пациента', LastName AS 'Фамилия', Name AS 'Имя', Patronymic AS 'Отчество', Pol AS 'Пол', Phone AS 'Телефон', Birthday AS 'Дата рождения', Medical_card AS 'Медицинская карта' From Patient;", sqlConnection);
                    DataSet dataSet = new DataSet();     //создаем датасет для результатов выборки
                    //Запишем данные в таблицу формы
                    dataAdapter.Fill(dataSet);          //заполняем датасет с помощью адаптера
                    dataGridView1.DataSource = dataSet.Tables[0];
                    Table.tableName = "Table_Patient";
                    tabPage1.Text = "Пациенты";
                }

                //Код причины 2
                using (SqlConnection sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["HospitalLists"].ConnectionString))
                {
                    sqlConnection.Open();
                    SqlDataAdapter dataAdapter = new SqlDataAdapter("SELECT ID AS 'Код опекуна', LastName AS 'Фамилия', Name AS 'Имя', Patronymic AS 'Отчество', Description AS 'Статус' From PatientCare p JOIN Care_Code_Table c ON c.Code = p.Code_comm;", sqlConnection);
                    DataSet dataSet = new DataSet();     //создаем датасет для результатов выборки
                    //Запишем данные в таблицу формы
                    dataAdapter.Fill(dataSet);          //заполняем датасет с помощью адаптера
                    dataGridView2.DataSource = dataSet.Tables[0];
                    Table.tableName = "Table_PatientCare";
                    tabPage2.Text = "Опекуны";
                }
                resizecolumns(); //автоподбор ширины столбцов
                // Воспроизводим звук
                sound1.Play();
            }
        }

        private void историяПриемовToolStripMenuItem_Click(object sender, EventArgs e)
        {
            tabPagesRemoveAdd(1);
            SingletonClass autorizObject = SingletonClass.getInstance();
            //История приемов
            using (SqlConnection sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["HospitalLists"].ConnectionString))
            {
                sqlConnection.Open();
                SqlDataAdapter dataAdapter = new SqlDataAdapter("SELECT h.ID AS 'ИД Приема', Doctor_ID AS 'ID Доктора', Patient_ID AS 'ID Пациента', LastName AS 'Фамилия', Name AS 'Имя', Patronymic AS 'Отчество', DatePriem AS 'Дата Приема', Bolnichn AS 'Больничный да/нет', DescriptionSick AS 'Описание болезни', Medicaments AS 'Назначенные препараты', DateBolnNextPriem AS 'Следующая дата приема' From ReceptionHistory h JOIN Patient p ON p.ID = h.Patient_ID", sqlConnection);
                DataSet dataSet = new DataSet();     //создаем датасет для результатов выборки
                                                     //Запишем данные в таблицу формы
                dataAdapter.Fill(dataSet);          //заполняем датасет с помощью адаптера
                dataGridView1.DataSource = dataSet.Tables[0];
                Table.tableName = "Table_ReceptionHistory";
                tabPage1.Text = "История приемов";
            }
            resizecolumns(); //автоподбор ширины столбцов
            // Воспроизводим звук
            sound1.Play();
        }

        private void больничныеЛистыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            tabPagesRemoveAdd(1);
            больничныйЛистToolStripMenuItem.Enabled = true;
            больничныйЛистToolStripMenuItem.Visible = true;
            SingletonClass autorizObject = SingletonClass.getInstance();
            if (autorizObject.getField1() == 1 || autorizObject.getField1() == 2) //если авторизован руководитель или врач
            {
                //Больничный лист
                using (SqlConnection sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["HospitalLists"].ConnectionString))
                {
                    sqlConnection.Open();
                    SqlDataAdapter dataAdapter = new SqlDataAdapter("SELECT b.Code AS '№ больничного', b.Patient_ID AS 'ID Пациента', p.LastName AS 'Фамилия', p.Name AS 'Имя', p.Patronymic AS 'Отчество', Pervich AS 'Первичный лист да/нет', b.Opecun_ID AS 'ИД опекуна', r1.Description AS 'Основная причина',r2.Description AS 'Дополнительная причина', Date_Issue AS 'Дата закрытия', StationarDateStart AS 'Дата нач. стационара', StationarDateEnd AS 'Дата оконч. стационара', Pregnancy AS 'Беременность', Distrubance_date AS 'Дата Нарушения', Distrubance_code AS 'Код Нарушения', MSE_Buro_Date AS 'Дата регистрации для бюро МСЕ', StartWork AS 'Дата начала работы', Inoe_Code AS 'Иное', History_ID AS 'ИД истории' From Hospital_List b JOIN Patient p ON p.ID = b.Patient_ID JOIN PatientCare c ON c.ID = b.Opecun_ID JOIN Reason_1 r1 ON r1.ID = b.Reason_Code1 JOIN Rereason_2 r2 ON r2.Code = b.Reason_Code2", sqlConnection);
                    DataSet dataSet = new DataSet();     //создаем датасет для результатов выборки
                    //Запишем данные в таблицу формы
                    dataAdapter.Fill(dataSet);          //заполняем датасет с помощью адаптера
                    dataGridView1.DataSource = dataSet.Tables[0];
                    Table.tableName = "Table_Hospital_List";
                    tabPage1.Text = "Больничные листы";
                }
            }
            resizecolumns(); //автоподбор ширины столбцов
            // Воспроизводим звук
            sound1.Play();
        }

        //Метод проверки что данные выбраны и не пусты
        private bool GetSelectedRowData(DataGridView dataGridView, List<string> dataArr)
        {
            if (dataGridView.SelectedRows.Count != 1)
            {
                MessageBox.Show("Выберите одну строку!", "Внимание!");
                return false;
            }
            //запоминаем выбранную строку
            int index = dataGridView.SelectedRows[0].Index;
            //Проверяем даннные в таблице
            for (int i = 0; i < dataGridView.Columns.Count; i++)
            {
                if (dataGridView.Rows[index].Cells[i].Value == null)
                {
                    MessageBox.Show("Не все данные введены!", "Внимание!");
                    return false;
                }
                dataArr.Add(dataGridView.Rows[index].Cells[i].Value.ToString());
            }
            return true;
        }

        //Метод отправки запроса с проверкой корректности
        private void ExecuteQuery(string query, SqlParameter[] parameters, SqlConnection sqlConnection)
        {
            try
            {
                SqlCommand command = new SqlCommand(query, sqlConnection);
                command.Parameters.AddRange(parameters);

                if (command.ExecuteNonQuery() != 1)
                    MessageBox.Show("Ошибка выполнения запроса!", "Ошибка!");
                else
                {
                    soundAdd.Play();
                    MessageBox.Show("Данные обработаны!", "Внимание!");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


        private void buttonAdd_Click(object sender, EventArgs e)
        {
            List<string> dataArr2 = new List<string>();
            string query = "";  //запрос


            //Учреждения
            if (tabControl1.SelectedIndex == 0 && tabPage1.Text == "Учреждения")
            {
                if (!GetSelectedRowData(dataGridView1, dataArr2))
                {
                    return;
                }

                // Создаем параметры запроса
                SqlParameter[] parameters = new SqlParameter[]
                {
                    new SqlParameter("@MedicalName", dataArr2[1]),
                    new SqlParameter("@Adress", dataArr2[2]),
                    new SqlParameter("@ORGN", dataArr2[3])
                };

                // Создаем запрос
                query = "INSERT INTO Medical_Institution (MedicalName, Adress, ORGN) VALUES (@MedicalName, @Adress, @ORGN)";

                // Проверяем и отправляем запрос
                using (SqlConnection sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["HospitalLists"].ConnectionString))
                {
                    sqlConnection.Open();
                    ExecuteQuery(query, parameters, sqlConnection); //отправляем запрос и проверяем данные
                }
            }

            //Отделения
            else if (tabControl1.SelectedIndex == 1 && tabPage2.Text == "Отделения")
            {
                if (!GetSelectedRowData(dataGridView2, dataArr2))
                {
                    return;
                }

                // Создаем параметры запроса
                SqlParameter[] parameters = new SqlParameter[]
                {
                    new SqlParameter("@Department", dataArr2[1]),
                    new SqlParameter("@MedicalInst_ID", dataArr2[2]),
                };

                // Создаем запрос
                query = "INSERT INTO Department VALUES(@Department, @MedicalInst_ID)";
                MessageBox.Show("В столбец Название введите id учреждения, к которому относится отдел");
                // Проверяем и отправляем запрос
                using (SqlConnection sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["HospitalLists"].ConnectionString))
                {
                    sqlConnection.Open();
                    ExecuteQuery(query, parameters, sqlConnection); //отправляем запрос и проверяем данные
                }
            }

            //Должности
            else if (tabControl1.SelectedIndex == 2 && tabPage3.Text == "Должности")
            {
                if (!GetSelectedRowData(dataGridView3, dataArr2))
                    return;

                // Создаем параметры запроса
                SqlParameter[] parameters = new SqlParameter[]
                {
                    new SqlParameter("@PostName", dataArr2[1]),
                    new SqlParameter("@Department_ID", dataArr2[2]),
                };

                // Создаем запрос
                query = "INSERT INTO Post VALUES(@PostName, @Department_ID)";
                MessageBox.Show("В столбец 'Название отдела' введите id отдела, за которым закреплена должность");
                // Проверяем и отправляем запрос
                using (SqlConnection sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["HospitalLists"].ConnectionString))
                {
                    sqlConnection.Open();
                    ExecuteQuery(query, parameters, sqlConnection); //отправляем запрос и проверяем данные
                }
            }

            //Спец код1
            else if (tabControl1.SelectedIndex == 0 && tabPage1.Text == "Причина (основная)")
            {
                if (!GetSelectedRowData(dataGridView1, dataArr2))
                    return;

                // Создаем параметры запроса
                SqlParameter[] parameters = new SqlParameter[]
                {
                    new SqlParameter("@ID", dataArr2[0]),
                    new SqlParameter("@Description", dataArr2[1]),
                };

                // Создаем запрос
                query = "INSERT INTO Reason_1 VALUES(@ID, @Description)";
                // Проверяем и отправляем запрос
                using (SqlConnection sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["HospitalLists"].ConnectionString))
                {
                    sqlConnection.Open();
                    ExecuteQuery(query, parameters, sqlConnection); //отправляем запрос и проверяем данные
                }
            }

            //Спец код2
            else if (tabControl1.SelectedIndex == 1 && tabPage2.Text == "Причина (дополнительная)")
            {
                if (!GetSelectedRowData(dataGridView2, dataArr2))
                    return;

                // Создаем параметры запроса
                SqlParameter[] parameters = new SqlParameter[]
                {
                    new SqlParameter("@Code", dataArr2[0]),
                    new SqlParameter("@Description", dataArr2[1]),
                };

                // Создаем запрос
                query = "INSERT INTO Rereason_2 VALUES(@Code, @Description)";
                // Проверяем и отправляем запрос
                using (SqlConnection sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["HospitalLists"].ConnectionString))
                {
                    sqlConnection.Open();
                    ExecuteQuery(query, parameters, sqlConnection); //отправляем запрос и проверяем данные
                }
            }

            //Нарушение режима
            else if (tabControl1.SelectedIndex == 2 && tabPage3.Text == "Нарушение режима")
            {
                if (!GetSelectedRowData(dataGridView3, dataArr2))
                    return;

                // Создаем параметры запроса
                SqlParameter[] parameters = new SqlParameter[]
                {
                    new SqlParameter("@Code", dataArr2[0]),
                    new SqlParameter("@Description", dataArr2[1]),
                };

                // Создаем запрос
                query = "INSERT INTO DisturbanceRegime VALUES(@Code, @Description)";
                // Проверяем и отправляем запрос
                using (SqlConnection sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["HospitalLists"].ConnectionString))
                {
                    sqlConnection.Open();
                    ExecuteQuery(query, parameters, sqlConnection); //отправляем запрос и проверяем данные
                }
            }

            //Родственные связи
            else if (tabControl1.SelectedIndex == 3 && tabPage4.Text == "Родственные связи")
            {
                if (!GetSelectedRowData(dataGridView4, dataArr2))
                    return;

                // Создаем параметры запроса
                SqlParameter[] parameters = new SqlParameter[]
                {
                    new SqlParameter("@Code", dataArr2[0]),
                    new SqlParameter("@Description", dataArr2[1]),
                };

                // Создаем запрос
                query = "INSERT INTO Care_Code_Table VALUES(@Code, @Description)";
                // Проверяем и отправляем запрос
                using (SqlConnection sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["HospitalLists"].ConnectionString))
                {
                    sqlConnection.Open();
                    ExecuteQuery(query, parameters, sqlConnection); //отправляем запрос и проверяем данные
                }
            }

            //Иное
            else if (tabControl1.SelectedIndex == 4 && tabPage5.Text == "Иное")
            {
                if (!GetSelectedRowData(dataGridView5, dataArr2))
                    return;

                // Создаем параметры запроса
                SqlParameter[] parameters = new SqlParameter[]
                {
                    new SqlParameter("@Code", dataArr2[0]),
                    new SqlParameter("@Description", dataArr2[1]),
                };

                // Создаем запрос
                query = "INSERT INTO Inoe_Codes VALUES(@Code, @Description)";
                // Проверяем и отправляем запрос
                using (SqlConnection sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["HospitalLists"].ConnectionString))
                {
                    sqlConnection.Open();
                    ExecuteQuery(query, parameters, sqlConnection); //отправляем запрос и проверяем данные
                }
            }

            //Врачи
            else if (tabControl1.SelectedIndex == 0 && tabPage1.Text == "Врачи")
            {
                if (!GetSelectedRowData(dataGridView1, dataArr2))
                    return;

                // Создаем параметры запроса
                SqlParameter[] parameters = new SqlParameter[]
                {
                    new SqlParameter("@LastName", dataArr2[1]),
                    new SqlParameter("@Name", dataArr2[2]),
                    new SqlParameter("@Patronymic", dataArr2[3]),
                    new SqlParameter("@Phone", dataArr2[4]),
                    new SqlParameter("@Post_ID", dataArr2[5]),
                };

                // Создаем запрос
                query = "INSERT INTO Doctor VALUES(@LastName, @Name, @Patronymic, @Phone, @Post_ID)";
                MessageBox.Show("В столбец Должность введите id должности специалиста");
                // Проверяем и отправляем запрос
                using (SqlConnection sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["HospitalLists"].ConnectionString))
                {
                    sqlConnection.Open();
                    ExecuteQuery(query, parameters, sqlConnection); //отправляем запрос и проверяем данные
                }
            }

            //Пациенты
            else if (tabControl1.SelectedIndex == 0 && tabPage1.Text == "Пациенты")
            {
                if (!GetSelectedRowData(dataGridView1, dataArr2))
                    return;

                // Создаем параметры запроса
                SqlParameter[] parameters = new SqlParameter[]
                {
                    new SqlParameter("@LastName", dataArr2[1]),
                    new SqlParameter("@Name", dataArr2[2]),
                    new SqlParameter("@Patronymic", dataArr2[3]),
                    new SqlParameter("@Pol", dataArr2[4]),
                    new SqlParameter("@Phone", dataArr2[5]),
                    new SqlParameter("@Birthday", dataArr2[6]),
                    new SqlParameter("@Medical_card", dataArr2[7])
                };

                // Создаем запрос
                query = "INSERT INTO Patient VALUES(@LastName, @Name, @Patronymic, @Pol, @Phone, @Birthday, @Medical_card)";
                // Проверяем и отправляем запрос
                using (SqlConnection sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["HospitalLists"].ConnectionString))
                {
                    sqlConnection.Open();
                    ExecuteQuery(query, parameters, sqlConnection); //отправляем запрос и проверяем данные
                }
            }

            //Опекуны
            else if (tabControl1.SelectedIndex == 1 && tabPage2.Text == "Опекуны")
            {
                if (!GetSelectedRowData(dataGridView2, dataArr2))
                    return;

                // Создаем параметры запроса
                SqlParameter[] parameters = new SqlParameter[]
                {
                    new SqlParameter("@LastName", dataArr2[1]),
                    new SqlParameter("@Name", dataArr2[2]),
                    new SqlParameter("@Patronymic", dataArr2[3]),
                    new SqlParameter("@Code_comm", dataArr2[4])
                };

                // Создаем запрос
                query = "INSERT INTO PatientCare VALUES(@LastName, @Name, @Patronymic, @Code_comm)";
                // Проверяем и отправляем запрос
                using (SqlConnection sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["HospitalLists"].ConnectionString))
                {
                    sqlConnection.Open();
                    ExecuteQuery(query, parameters, sqlConnection); //отправляем запрос и проверяем данные
                }
            }

            /*
            //История приемов
            else if (tabControl1.SelectedIndex == 0 && tabPage2.Text == "История приемов")
            {
                if (!GetSelectedRowData(dataGridView1, dataArr2))
                    return;

                // Создаем параметры запроса
                SqlParameter[] parameters = new SqlParameter[]
                {
                    new SqlParameter("@LastName", dataArr2[1]),
                    new SqlParameter("@Name", dataArr2[2]),
                    new SqlParameter("@Patronymic", dataArr2[3]),
                    new SqlParameter("@Code_comm", dataArr2[4])
                };

                // Создаем запрос
                query = "INSERT INTO PatientCare VALUES(@LastName, @Name, @Patronymic, @Code_comm)";
                // Проверяем и отправляем запрос
                using (SqlConnection sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["HospitalLists"].ConnectionString))
                {
                    sqlConnection.Open();
                    ExecuteQuery(query, parameters, sqlConnection); //отправляем запрос и проверяем данные
                }
            }
            */

            //История приемов
            else if (tabControl1.SelectedIndex == 0 && tabPage1.Text == "История приемов")
            {
                if (!GetSelectedRowData(dataGridView1, dataArr2))
                    return;

                DateTime datePriem = DateTime.Parse(dataArr2[6]);
                DateTime dateBolnNextPriem = DateTime.Parse(dataArr2[10]);
                // Создаем параметры запроса
                SqlParameter[] parameters = new SqlParameter[]
                {
                    new SqlParameter("@Patient_ID", dataArr2[1]),
                    new SqlParameter("@Doctor_ID", dataArr2[2]),
                    new SqlParameter("@DatePriem", datePriem),
                    new SqlParameter("@Bolnichn", dataArr2[7]),
                    new SqlParameter("@DescriptionSick", dataArr2[8]),
                    new SqlParameter("@Medicaments", dataArr2[9]),
                    new SqlParameter("@DateBolnNextPriem", dateBolnNextPriem),
                };

                MessageBox.Show("в столбец 'Пациент' введите id Пациента, В столбец 'Доктор' введите id специалиста");
                // Создаем запрос
                query = "INSERT INTO ReceptionHistory VALUES(@Patient_ID, @Doctor_ID, @DatePriem, @Bolnichn, @DescriptionSick, @Medicaments, @DateBolnNextPriem)";
                // Проверяем и отправляем запрос
                using (SqlConnection sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["HospitalLists"].ConnectionString))
                {
                    sqlConnection.Open();
                    ExecuteQuery(query, parameters, sqlConnection); //отправляем запрос и проверяем данные
                }
            }

            //Больничные листы
            else if (tabControl1.SelectedIndex == 0 && tabPage1.Text == "Больничные листы")
            {
                if (!GetSelectedRowData(dataGridView1, dataArr2))
                    return;


                DateTime dateIssue = DateTime.Parse(dataArr2[9]);
                DateTime dateStationStart = DateTime.Parse(dataArr2[10]);
                DateTime dateStationEnd = DateTime.Parse(dataArr2[11]);
                DateTime Distrubance_date = DateTime.Parse(dataArr2[13]);
                DateTime MSE_Buro_Date = DateTime.Parse(dataArr2[15]);
                DateTime dateStartWork = DateTime.Parse(dataArr2[16]);

                // Создаем параметры запроса
                SqlParameter[] parameters = new SqlParameter[]
                {
                    new SqlParameter("@Code", dataArr2[0]),
                    new SqlParameter("@Pervich", dataArr2[5]),
                    new SqlParameter("@Date_Issue", dateIssue),
                    new SqlParameter("@Patient_ID", dataArr2[1]),
                    new SqlParameter("@Opecun_ID", dataArr2[6]),
                    new SqlParameter("@Reason_Code1", dataArr2[7]),
                    new SqlParameter("@Reason_Code2", dataArr2[8]),
                    new SqlParameter("@StationarDateStart", dateStationStart),
                    new SqlParameter("@StationarDateEnd", dateStationEnd),
                    new SqlParameter("@Pregnancy", dataArr2[12]),
                    new SqlParameter("@Distrubance_date", Distrubance_date),
                    new SqlParameter("@Distrubance_code", dataArr2[14]),
                    new SqlParameter("@MSE_Buro_Date", MSE_Buro_Date),
                    new SqlParameter("@StartWork", dateStartWork),
                    new SqlParameter("@Inoe_Code", dataArr2[17]),
                    new SqlParameter("@History_ID", dataArr2[18]),
                };

                //MessageBox.Show("В столбец 'Пациент' введите id Пациента, В столбец 'Доктор' введите id специалиста");
                // Создаем запрос
                query = "INSERT INTO Hospital_List VALUES(@Code, @Pervich, @Date_Issue, @Patient_ID, @Opecun_ID, @Reason_Code1, @Reason_Code2, @StationarDateStart, @StationarDateEnd, @Pregnancy, @Distrubance_date, @Distrubance_code, @MSE_Buro_Date, @StartWork, @Inoe_Code, @History_ID)";
                // Проверяем и отправляем запрос
                using (SqlConnection sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["HospitalLists"].ConnectionString))
                {
                    sqlConnection.Open();
                    ExecuteQuery(query, parameters, sqlConnection); //отправляем запрос и проверяем данные
                }
            }

        }

        //Учреждения UPDATE
        private void buttonUpdate_Click(object sender, EventArgs e)
        {
            List<string> dataArr2 = new List<string>();
            string query = "";  //запрос


            //Учреждения
            if (tabControl1.SelectedIndex == 0 && tabPage1.Text == "Учреждения")
            {
                if (!GetSelectedRowData(dataGridView1, dataArr2))
                {
                    return;
                }

                // Создаем параметры запроса
                SqlParameter[] parameters = new SqlParameter[]
                {
                    new SqlParameter("@ID", dataArr2[0]),
                    new SqlParameter("@MedicalName", dataArr2[1]),
                    new SqlParameter("@Adress", dataArr2[2]),
                    new SqlParameter("@ORGN", dataArr2[3])
                };

                // Создаем запрос
                query = "UPDATE Medical_Institution SET MedicalName = @MedicalName, Adress = @Adress WHERE ID = @ID";

                // Проверяем и отправляем запрос
                using (SqlConnection sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["HospitalLists"].ConnectionString))
                {
                    sqlConnection.Open();
                    ExecuteQuery(query, parameters, sqlConnection); //отправляем запрос и проверяем данные
                }
            }

            //Отделения
            else if (tabControl1.SelectedIndex == 1 && tabPage2.Text == "Отделения")
            {
                if (!GetSelectedRowData(dataGridView2, dataArr2))
                {
                    return;
                }

                // Создаем параметры запроса
                SqlParameter[] parameters = new SqlParameter[]
                {
                    new SqlParameter("@ID", dataArr2[0]),
                    new SqlParameter("@Department", dataArr2[1]),
                    new SqlParameter("@MedicalInst_ID", dataArr2[2]),
                };

                // Создаем запрос
                query = "UPDATE Department SET Department = @Department, MedicalInst_ID = @MedicalInst_ID WHERE ID = @ID";
                MessageBox.Show("В столбец Название введите id учреждения, к которому относится отдел");
                // Проверяем и отправляем запрос
                using (SqlConnection sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["HospitalLists"].ConnectionString))
                {
                    sqlConnection.Open();
                    ExecuteQuery(query, parameters, sqlConnection); //отправляем запрос и проверяем данные
                }
            }

            //Должности
            else if (tabControl1.SelectedIndex == 2 && tabPage3.Text == "Должности")
            {
                if (!GetSelectedRowData(dataGridView3, dataArr2))
                    return;

                // Создаем параметры запроса
                SqlParameter[] parameters = new SqlParameter[]
                {
                    new SqlParameter("@ID", dataArr2[0]),
                    new SqlParameter("@PostName", dataArr2[1]),
                    new SqlParameter("@Department_ID", dataArr2[2])
                };

                // Создаем запрос
                query = "UPDATE Post SET PostName = @PostName, Department_ID = @Department_ID WHERE ID = @ID";
                MessageBox.Show("В столбец 'Название отдела' введите id отдела, за которым закреплена должность");
                // Проверяем и отправляем запрос
                using (SqlConnection sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["HospitalLists"].ConnectionString))
                {
                    sqlConnection.Open();
                    ExecuteQuery(query, parameters, sqlConnection); //отправляем запрос и проверяем данные
                }
            }

            //Спец код1
            else if (tabControl1.SelectedIndex == 0 && tabPage1.Text == "Причина (основная)")
            {
                if (!GetSelectedRowData(dataGridView1, dataArr2))
                    return;

                // Создаем параметры запроса
                SqlParameter[] parameters = new SqlParameter[]
                {
                    new SqlParameter("@ID", dataArr2[0]),
                    new SqlParameter("@Description", dataArr2[1]),
                };

                // Создаем запрос
                query = "UPDATE Reason_1 SET Description = @Description WHERE ID = @ID";
                // Проверяем и отправляем запрос
                using (SqlConnection sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["HospitalLists"].ConnectionString))
                {
                    sqlConnection.Open();
                    ExecuteQuery(query, parameters, sqlConnection); //отправляем запрос и проверяем данные
                }
            }

            //Спец код2
            else if (tabControl1.SelectedIndex == 1 && tabPage2.Text == "Причина (дополнительная)")
            {
                if (!GetSelectedRowData(dataGridView2, dataArr2))
                    return;

                // Создаем параметры запроса
                SqlParameter[] parameters = new SqlParameter[]
                {
                    new SqlParameter("@Code", dataArr2[0]),
                    new SqlParameter("@Description", dataArr2[1]),
                };

                // Создаем запрос
                query = "UPDATE Rereason_2 SET Description = @Description WHERE Code = @Code";
                // Проверяем и отправляем запрос
                using (SqlConnection sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["HospitalLists"].ConnectionString))
                {
                    sqlConnection.Open();
                    ExecuteQuery(query, parameters, sqlConnection); //отправляем запрос и проверяем данные
                }
            }

            //Нарушение режима
            else if (tabControl1.SelectedIndex == 2 && tabPage3.Text == "Нарушение режима")
            {
                if (!GetSelectedRowData(dataGridView3, dataArr2))
                    return;

                // Создаем параметры запроса
                SqlParameter[] parameters = new SqlParameter[]
                {
                    new SqlParameter("@Code", dataArr2[0]),
                    new SqlParameter("@Description", dataArr2[1]),
                };

                // Создаем запрос
                query = "UPDATE DisturbanceRegime SET Description = @Description WHERE Code = @Code";
                // Проверяем и отправляем запрос
                using (SqlConnection sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["HospitalLists"].ConnectionString))
                {
                    sqlConnection.Open();
                    ExecuteQuery(query, parameters, sqlConnection); //отправляем запрос и проверяем данные
                }
            }

            //Родственные связи
            else if (tabControl1.SelectedIndex == 3 && tabPage4.Text == "Родственные связи")
            {
                if (!GetSelectedRowData(dataGridView4, dataArr2))
                    return;

                // Создаем параметры запроса
                SqlParameter[] parameters = new SqlParameter[]
                {
                    new SqlParameter("@Code", dataArr2[0]),
                    new SqlParameter("@Description", dataArr2[1]),
                };

                // Создаем запрос
                query = "UPDATE Care_Code_Table SET Description = @Description WHERE Code = @Code";
                // Проверяем и отправляем запрос
                using (SqlConnection sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["HospitalLists"].ConnectionString))
                {
                    sqlConnection.Open();
                    ExecuteQuery(query, parameters, sqlConnection); //отправляем запрос и проверяем данные
                }
            }

            //Иное
            else if (tabControl1.SelectedIndex == 4 && tabPage5.Text == "Иное")
            {
                if (!GetSelectedRowData(dataGridView5, dataArr2))
                    return;

                // Создаем параметры запроса
                SqlParameter[] parameters = new SqlParameter[]
                {
                    new SqlParameter("@Code", dataArr2[0]),
                    new SqlParameter("@Description", dataArr2[1]),
                };

                // Создаем запрос
                query = "UPDATE Inoe_Codes SET Description = @Description WHERE Code = @Code";
                // Проверяем и отправляем запрос
                using (SqlConnection sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["HospitalLists"].ConnectionString))
                {
                    sqlConnection.Open();
                    ExecuteQuery(query, parameters, sqlConnection); //отправляем запрос и проверяем данные
                }
            }

            //Врачи
            else if (tabControl1.SelectedIndex == 0 && tabPage1.Text == "Врачи")
            {
                if (!GetSelectedRowData(dataGridView1, dataArr2))
                    return;

                // Создаем параметры запроса
                SqlParameter[] parameters = new SqlParameter[]
                {
                    new SqlParameter("@ID", dataArr2[0]),
                    new SqlParameter("@LastName", dataArr2[1]),
                    new SqlParameter("@Name", dataArr2[2]),
                    new SqlParameter("@Patronymic", dataArr2[3]),
                    new SqlParameter("@Phone", dataArr2[4]),
                    new SqlParameter("@Post_ID", dataArr2[5]),
                };

                // Создаем запрос
                query = "UPDATE Doctor SET LastName = @LastName, Name = @Name, Patronymic = @Patronymic, Phone = @Phone, Post_ID = @Post_ID WHERE ID = @ID";
                MessageBox.Show("В столбец Должность введите id должности специалиста");
                // Проверяем и отправляем запрос
                using (SqlConnection sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["HospitalLists"].ConnectionString))
                {
                    sqlConnection.Open();
                    ExecuteQuery(query, parameters, sqlConnection); //отправляем запрос и проверяем данные
                }
            }

            //Пациенты
            else if (tabControl1.SelectedIndex == 0 && tabPage1.Text == "Пациенты")
            {
                if (!GetSelectedRowData(dataGridView1, dataArr2))
                    return;

                // Создаем параметры запроса
                SqlParameter[] parameters = new SqlParameter[]
                {
                    new SqlParameter("@ID", dataArr2[0]),
                    new SqlParameter("@LastName", dataArr2[1]),
                    new SqlParameter("@Name", dataArr2[2]),
                    new SqlParameter("@Patronymic", dataArr2[3]),
                    new SqlParameter("@Pol", dataArr2[4]),
                    new SqlParameter("@Phone", dataArr2[5]),
                    new SqlParameter("@Birthday", dataArr2[6]),
                    new SqlParameter("@Medical_card", dataArr2[7])
                };

                // Создаем запрос
                query = "UPDATE Patient SET LastName = @LastName, Name = @Name, Patronymic = @Patronymic, Pol = @Pol, Phone = @Phone, Birthday = @Birthday, Medical_card = @Medical_card WHERE ID = @ID";
                // Проверяем и отправляем запрос
                using (SqlConnection sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["HospitalLists"].ConnectionString))
                {
                    sqlConnection.Open();
                    ExecuteQuery(query, parameters, sqlConnection); //отправляем запрос и проверяем данные
                }
            }

            //Опекуны
            else if (tabControl1.SelectedIndex == 1 && tabPage2.Text == "Опекуны")
            {
                if (!GetSelectedRowData(dataGridView2, dataArr2))
                    return;

                // Создаем параметры запроса
                SqlParameter[] parameters = new SqlParameter[]
                {
                    new SqlParameter("@ID", dataArr2[0]),
                    new SqlParameter("@LastName", dataArr2[1]),
                    new SqlParameter("@Name", dataArr2[2]),
                    new SqlParameter("@Patronymic", dataArr2[3]),
                    new SqlParameter("@Code_comm", dataArr2[4])
                };

                // Создаем запрос
                query = "UPDATE PatientCare SET LastName = @LastName, Name = @Name, Patronymic = @Patronymic, Code_comm = @Code_comm WHERE ID = @ID";
                // Проверяем и отправляем запрос
                using (SqlConnection sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["HospitalLists"].ConnectionString))
                {
                    sqlConnection.Open();
                    ExecuteQuery(query, parameters, sqlConnection); //отправляем запрос и проверяем данные
                }
            }

            //История приемов
            else if (tabControl1.SelectedIndex == 0 && tabPage1.Text == "История приемов")
            {
                if (!GetSelectedRowData(dataGridView1, dataArr2))
                    return;

                DateTime datePriem = DateTime.Parse(dataArr2[6]);
                DateTime dateBolnNextPriem = DateTime.Parse(dataArr2[10]);
                // Создаем параметры запроса
                SqlParameter[] parameters = new SqlParameter[]
                {
                    new SqlParameter("@ID", dataArr2[0]),
                    new SqlParameter("@Patient_ID", dataArr2[2]),
                    new SqlParameter("@Doctor_ID", dataArr2[1]),
                    new SqlParameter("@DatePriem", datePriem),
                    new SqlParameter("@Bolnichn", dataArr2[7]),
                    new SqlParameter("@DescriptionSick", dataArr2[8]),
                    new SqlParameter("@Medicaments", dataArr2[9]),
                    new SqlParameter("@DateBolnNextPriem", dateBolnNextPriem),
                };

                MessageBox.Show("в столбец 'Пациент' введите id Пациента, В столбец 'Доктор' введите id специалиста");
                // Создаем запрос
                query = "UPDATE ReceptionHistory SET Patient_ID = @Patient_ID, Doctor_ID = @Doctor_ID, DatePriem = @DatePriem, Bolnichn = @Bolnichn, DescriptionSick = @DescriptionSick, Medicaments = @Medicaments, DateBolnNextPriem = @DateBolnNextPriem WHERE ID = @ID";
                // Проверяем и отправляем запрос
                using (SqlConnection sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["HospitalLists"].ConnectionString))
                {
                    sqlConnection.Open();
                    ExecuteQuery(query, parameters, sqlConnection); //отправляем запрос и проверяем данные
                }
            }

            //Больничные листы
            else if (tabControl1.SelectedIndex == 0 && tabPage1.Text == "Больничные листы")
            {
                MessageBox.Show("Запрещено редактирование больничных листов!\n" +
                    "Обратитесь к системному администратору!");
            }

        }

        private void buttonDelete_Click(object sender, EventArgs e)
        {
            List<string> dataArr2 = new List<string>();
            string query = "";  //запрос

            //Учреждения
            if (tabControl1.SelectedIndex == 0 && tabPage1.Text == "Учреждения")
            {
                if (!GetSelectedRowData(dataGridView1, dataArr2))
                {
                    return;
                }

                // Создаем параметры запроса
                SqlParameter[] parameters = new SqlParameter[]
                {
                    new SqlParameter("@ID", dataArr2[0]),
                };

                // Создаем запрос
                query = "DELETE FROM Medical_Institution WHERE ID = @ID";

                // Проверяем и отправляем запрос
                using (SqlConnection sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["HospitalLists"].ConnectionString))
                {
                    sqlConnection.Open();
                    ExecuteQuery(query, parameters, sqlConnection); //отправляем запрос и проверяем данные
                }
            }

            //Отделения
            else if (tabControl1.SelectedIndex == 1 && tabPage2.Text == "Отделения")
            {
                if (!GetSelectedRowData(dataGridView2, dataArr2))
                {
                    return;
                }

                // Создаем параметры запроса
                SqlParameter[] parameters = new SqlParameter[]
                {
                    new SqlParameter("@ID", dataArr2[0]),
                };

                // Создаем запрос
                query = "DELETE FROM Department WHERE ID = @ID";
                // Проверяем и отправляем запрос
                using (SqlConnection sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["HospitalLists"].ConnectionString))
                {
                    sqlConnection.Open();
                    ExecuteQuery(query, parameters, sqlConnection); //отправляем запрос и проверяем данные
                }
            }

            //Должности
            else if (tabControl1.SelectedIndex == 2 && tabPage3.Text == "Должности")
            {
                if (!GetSelectedRowData(dataGridView3, dataArr2))
                    return;

                // Создаем параметры запроса
                SqlParameter[] parameters = new SqlParameter[]
                {
                    new SqlParameter("@ID", dataArr2[0]),
                };

                // Создаем запрос
                query = "DELETE FROM Post WHERE ID = @ID";
                // Проверяем и отправляем запрос
                using (SqlConnection sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["HospitalLists"].ConnectionString))
                {
                    sqlConnection.Open();
                    ExecuteQuery(query, parameters, sqlConnection); //отправляем запрос и проверяем данные
                }
            }

            //Спец код1
            else if (tabControl1.SelectedIndex == 0 && tabPage1.Text == "Причина (основная)")
            {
                if (!GetSelectedRowData(dataGridView1, dataArr2))
                    return;

                // Создаем параметры запроса
                SqlParameter[] parameters = new SqlParameter[]
                {
                    new SqlParameter("@ID", dataArr2[0]),
                };

                // Создаем запрос
                query = "DELETE FROM Reason_1 WHERE ID = @ID";
                // Проверяем и отправляем запрос
                using (SqlConnection sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["HospitalLists"].ConnectionString))
                {
                    sqlConnection.Open();
                    ExecuteQuery(query, parameters, sqlConnection); //отправляем запрос и проверяем данные
                }
            }

            //Спец код2
            else if (tabControl1.SelectedIndex == 1 && tabPage2.Text == "Причина (дополнительная)")
            {
                if (!GetSelectedRowData(dataGridView2, dataArr2))
                    return;

                // Создаем параметры запроса
                SqlParameter[] parameters = new SqlParameter[]
                {
                    new SqlParameter("@Code", dataArr2[0]),
                };

                // Создаем запрос
                query = "DELETE FROM Rereason_2 WHERE Code = @Code";
                // Проверяем и отправляем запрос
                using (SqlConnection sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["HospitalLists"].ConnectionString))
                {
                    sqlConnection.Open();
                    ExecuteQuery(query, parameters, sqlConnection); //отправляем запрос и проверяем данные
                }
            }

            //Нарушение режима
            else if (tabControl1.SelectedIndex == 2 && tabPage3.Text == "Нарушение режима")
            {
                if (!GetSelectedRowData(dataGridView3, dataArr2))
                    return;

                // Создаем параметры запроса
                SqlParameter[] parameters = new SqlParameter[]
                {
                    new SqlParameter("@Code", dataArr2[0])
                };

                // Создаем запрос
                query = "DELETE FROM DisturbanceRegime WHERE Code = @Code";
                // Проверяем и отправляем запрос
                using (SqlConnection sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["HospitalLists"].ConnectionString))
                {
                    sqlConnection.Open();
                    ExecuteQuery(query, parameters, sqlConnection); //отправляем запрос и проверяем данные
                }
            }

            //Родственные связи
            else if (tabControl1.SelectedIndex == 3 && tabPage4.Text == "Родственные связи")
            {
                if (!GetSelectedRowData(dataGridView4, dataArr2))
                    return;

                // Создаем параметры запроса
                SqlParameter[] parameters = new SqlParameter[]
                {
                    new SqlParameter("@Code", dataArr2[0])
                };

                // Создаем запрос
                query = "DELETE FROM Care_Code_Table WHERE Code = @Code";
                // Проверяем и отправляем запрос
                using (SqlConnection sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["HospitalLists"].ConnectionString))
                {
                    sqlConnection.Open();
                    ExecuteQuery(query, parameters, sqlConnection); //отправляем запрос и проверяем данные
                }
            }

            //Иное
            else if (tabControl1.SelectedIndex == 4 && tabPage5.Text == "Иное")
            {
                if (!GetSelectedRowData(dataGridView5, dataArr2))
                    return;

                // Создаем параметры запроса
                SqlParameter[] parameters = new SqlParameter[]
                {
                    new SqlParameter("@Code", dataArr2[0])
                };

                // Создаем запрос
                query = "DELETE FROM Inoe_Codes WHERE Code = @Code";
                // Проверяем и отправляем запрос
                using (SqlConnection sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["HospitalLists"].ConnectionString))
                {
                    sqlConnection.Open();
                    ExecuteQuery(query, parameters, sqlConnection); //отправляем запрос и проверяем данные
                }
            }

            //Врачи
            else if (tabControl1.SelectedIndex == 0 && tabPage1.Text == "Врачи")
            {
                if (!GetSelectedRowData(dataGridView1, dataArr2))
                    return;

                // Создаем параметры запроса
                SqlParameter[] parameters = new SqlParameter[]
                {
                    new SqlParameter("@ID", dataArr2[0])
                };

                // Создаем запрос
                query = "DELETE FROM Doctor WHERE ID = @ID";
                // Проверяем и отправляем запрос
                using (SqlConnection sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["HospitalLists"].ConnectionString))
                {
                    sqlConnection.Open();
                    ExecuteQuery(query, parameters, sqlConnection); //отправляем запрос и проверяем данные
                }
            }

            //Пациенты
            else if (tabControl1.SelectedIndex == 0 && tabPage1.Text == "Пациенты")
            {
                if (!GetSelectedRowData(dataGridView1, dataArr2))
                    return;

                // Создаем параметры запроса
                SqlParameter[] parameters = new SqlParameter[]
                {
                    new SqlParameter("@ID", dataArr2[0]),
                };

                // Создаем запрос
                query = "DELETE FROM Patient WHERE ID = @ID";
                // Проверяем и отправляем запрос
                using (SqlConnection sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["HospitalLists"].ConnectionString))
                {
                    sqlConnection.Open();
                    ExecuteQuery(query, parameters, sqlConnection); //отправляем запрос и проверяем данные
                }
            }

            //Опекуны
            else if (tabControl1.SelectedIndex == 1 && tabPage2.Text == "Опекуны")
            {
                if (!GetSelectedRowData(dataGridView2, dataArr2))
                    return;

                // Создаем параметры запроса
                SqlParameter[] parameters = new SqlParameter[]
                {
                    new SqlParameter("@ID", dataArr2[0])
                };

                // Создаем запрос
                query = "DELETE FROM PatientCare WHERE ID = @ID";
                // Проверяем и отправляем запрос
                using (SqlConnection sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["HospitalLists"].ConnectionString))
                {
                    sqlConnection.Open();
                    ExecuteQuery(query, parameters, sqlConnection); //отправляем запрос и проверяем данные
                }
            }

            //История приемов
            else if (tabControl1.SelectedIndex == 0 && tabPage1.Text == "История приемов")
            {
                if (!GetSelectedRowData(dataGridView1, dataArr2))
                    return;

                // Создаем параметры запроса
                SqlParameter[] parameters = new SqlParameter[]
                {
                    new SqlParameter("@ID", dataArr2[0])
                };
                // Создаем запрос
                query = "DELETE FROM ReceptionHistory WHERE ID = @ID";
                // Проверяем и отправляем запрос
                using (SqlConnection sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["HospitalLists"].ConnectionString))
                {
                    sqlConnection.Open();
                    ExecuteQuery(query, parameters, sqlConnection); //отправляем запрос и проверяем данные
                }
            }
            //Больничные листы
            else if (tabControl1.SelectedIndex == 0 && tabPage1.Text == "Больничные листы")
            {
                MessageBox.Show("Запрещено удаление больничных листов!\n" +
                    "Обратитесь к системному администратору!");
            }

        }

        private void справкаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Программа имеет несколько уровней доступа. " +
                "Руководитель может просматривать и изменять все таблицы. Врач может просматривать и изменять все таблицы, связанные с пациентами и больничными листами." +
                "Медсестра может просматривать назначенное лечение больному и препараты, а так же редактировать таблицу Приемов пациентов\n" +
                "Для добавления, изменения или удаления данных ячейку нужно выделить, заполнить необходимыми данными или отредактировать существующую и нажать на одну из управляющих кнопок", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void оПрограммеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("v 1.0.0", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void сToolStripMenuItem_Click(object sender, EventArgs e)
        {
            List<string> dataArr2 = new List<string>();
            string query = "";  //запрос
            if (!GetSelectedRowData(dataGridView1, dataArr2))
                return;
            try
            {
                // Создаем новый экземпляр приложения Excel
                Excel.Application excelApp = new Excel.Application();

                // Открываем книгу Excel
                string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "HospitalList.xlsx");

                // Проверяем существует ли файл
                if (!File.Exists(filePath))
                {
                    // Создаем новую книгу Excel
                    Excel.Workbook workbook = excelApp.Workbooks.Add();
                    workbook.SaveAs(filePath);
                    workbook.Close();
                }

                // Открываем существующую книгу Excel
                Excel.Workbook existingWorkbook = excelApp.Workbooks.Open(filePath);

                // Получаем ссылку на активный лист
                Excel.Worksheet worksheet = existingWorkbook.ActiveSheet;

                // Изменяем ширину первого столбца (A)
                worksheet.Columns[1].ColumnWidth = 30; // Устанавливаем ширину столбца A равной 20
                worksheet.Columns[2].ColumnWidth = 40;
                
                // Записываем данные в определенную ячейку
                worksheet.Cells[1, 1] = "№ больничного листа: "; // Записываем текстовое описание в указанную строку в столбце A
                worksheet.Cells[1, 2] = dataArr2[0]; // Записываем данные в указанную строку в столбце A
                worksheet.Cells[2, 1] = "Фамилия Имя Отчетво: ";
                worksheet.Cells[2, 2] = dataArr2[2]+ " " + dataArr2[3] + " " + dataArr2[4]; // Записываем данные в указанную строку в столбце A
                worksheet.Cells[3, 1] = "Первичный лист: ";
                worksheet.Cells[3, 2] = dataArr2[5];
                worksheet.Cells[4, 1] = "Основная причина: ";
                worksheet.Cells[4, 2] = dataArr2[7];
                worksheet.Cells[5, 1] = "Дополнительная причина: ";
                worksheet.Cells[5, 2] = dataArr2[8];
                worksheet.Cells[6, 1] = "Дата закрытия б/л: ";
                worksheet.Cells[6, 2] = dataArr2[9];
                worksheet.Cells[7, 1] = "Дата начала стационарного лечения: ";
                worksheet.Cells[7, 2] = dataArr2[10];
                worksheet.Cells[8, 1] = "Дата окончания стационарного лечения: ";
                worksheet.Cells[8, 2] = dataArr2[11];
                worksheet.Cells[9, 1] = "Беременость: ";
                worksheet.Cells[9, 2] = dataArr2[12];
                worksheet.Cells[10, 1] = "Нарушения режима: ";
                worksheet.Cells[10, 2] = dataArr2[13];
                worksheet.Cells[11, 1] = "Дата нарушения: ";
                worksheet.Cells[11, 2] = dataArr2[14];
                worksheet.Cells[12, 1] = "Дата регистрации для бюро МСЕ: ";
                worksheet.Cells[12, 2] = dataArr2[15];
                worksheet.Cells[13, 1] = "Дата начала работы: ";
                worksheet.Cells[13, 2] = dataArr2[16];
                worksheet.Cells[14, 1] = "Иное Код: ";
                worksheet.Cells[14, 2] = dataArr2[17];

                // Сохраняем изменения и закрываем книгу Excel
                existingWorkbook.Save();
                existingWorkbook.Close();

                // Закрываем приложение Excel
                excelApp.Quit();
                soundDrum.Play();
                MessageBox.Show("Файл успешно записан!");

                // Освобождаем ресурсы Excel
                if (existingWorkbook != null)
                    Marshal.ReleaseComObject(existingWorkbook);
                if (excelApp != null)
                    Marshal.ReleaseComObject(excelApp);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при открытии книги Excel: " + ex.Message);
            }
        }

        private void распечататьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "HospitalList.xlsx");
            // Создаем процесс для вызова команды печати
            ProcessStartInfo psi = new ProcessStartInfo
            {
                FileName = filePath,
                Verb = "Print"
            };

            // Запускаем процесс
            Process.Start(psi);
            soundPush.Play();
        }

        private void отправитьПоПочтеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "HospitalList.xlsx");
            MailMessage mail = new MailMessage();
            SmtpClient SmtpServer = new SmtpClient("smtp.gmail.com");

            mail.From = new MailAddress("your_email@gmail.com");
            mail.To.Add("recipient_email@gmail.com");
            mail.Subject = "Hospital List";
            mail.Body = "Please find attached the Hospital List.";

            System.Net.Mail.Attachment attachment;
            attachment = new System.Net.Mail.Attachment(filePath);
            mail.Attachments.Add(attachment);

            SmtpServer.Port = 587;
            SmtpServer.Credentials = new System.Net.NetworkCredential("your_email@gmail.com", "your_password");
            SmtpServer.EnableSsl = true;

            SmtpServer.Send(mail);

            MessageBox.Show("Email sent successfully.");
            soundWHOOSH.Play();
        }
    }



}
