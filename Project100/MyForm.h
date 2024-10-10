#pragma once

namespace Project100 {

    using namespace System;
    using namespace System::ComponentModel;
    using namespace System::Collections;
    using namespace System::Windows::Forms;
    using namespace System::Data;
    using namespace System::Drawing;
    using namespace System::Data::OleDb;

    public ref class MyForm : public System::Windows::Forms::Form
    {
    public:
        MyForm(void)
        {
            InitializeComponent();
        }

    protected:
        ~MyForm()
        {
            if (components)
            {
                delete components;
            }
        }

    private: System::Windows::Forms::MenuStrip^ menuStrip1;
    private: System::Windows::Forms::ToolStripMenuItem^ ä³¿ToolStripMenuItem;
    private: System::Windows::Forms::ToolStripMenuItem^ çàâàíòàæèòèToolStripMenuItem;
    private: System::Windows::Forms::ToolStripMenuItem^ ïğîÏğîãğàìóToolStripMenuItem;
    private: System::Windows::Forms::ToolStripMenuItem^ âèõ³äToolStripMenuItem;
    private: System::Windows::Forms::DataGridView^ dataGridView1;
    private: System::Windows::Forms::ToolStripMenuItem^ îíToolStripMenuItem;
    private: System::Windows::Forms::ToolStripMenuItem^ äîäàòèToolStripMenuItem;
    private: System::Windows::Forms::ToolStripMenuItem^ âèäàëèòèToolStripMenuItem;
    private: System::Windows::Forms::DataGridViewTextBoxColumn^ Id;
    private: System::Windows::Forms::DataGridViewTextBoxColumn^ ServiceName;
    private: System::Windows::Forms::DataGridViewTextBoxColumn^ Price;
    private: System::Windows::Forms::DataGridViewTextBoxColumn^ Deadline;
    private: System::Windows::Forms::DataGridViewTextBoxColumn^ Detail;

    private:
        System::ComponentModel::Container^ components;

#pragma region Windows Form Designer generated code
        void InitializeComponent(void)
        {
            this->menuStrip1 = (gcnew System::Windows::Forms::MenuStrip());
            this->ä³¿ToolStripMenuItem = (gcnew System::Windows::Forms::ToolStripMenuItem());
            this->çàâàíòàæèòèToolStripMenuItem = (gcnew System::Windows::Forms::ToolStripMenuItem());
            this->îíToolStripMenuItem = (gcnew System::Windows::Forms::ToolStripMenuItem());
            this->äîäàòèToolStripMenuItem = (gcnew System::Windows::Forms::ToolStripMenuItem());
            this->âèäàëèòèToolStripMenuItem = (gcnew System::Windows::Forms::ToolStripMenuItem());
            this->ïğîÏğîãğàìóToolStripMenuItem = (gcnew System::Windows::Forms::ToolStripMenuItem());
            this->âèõ³äToolStripMenuItem = (gcnew System::Windows::Forms::ToolStripMenuItem());
            this->dataGridView1 = (gcnew System::Windows::Forms::DataGridView());
            this->Id = (gcnew System::Windows::Forms::DataGridViewTextBoxColumn());
            this->ServiceName = (gcnew System::Windows::Forms::DataGridViewTextBoxColumn());
            this->Price = (gcnew System::Windows::Forms::DataGridViewTextBoxColumn());
            this->Deadline = (gcnew System::Windows::Forms::DataGridViewTextBoxColumn());
            this->Detail = (gcnew System::Windows::Forms::DataGridViewTextBoxColumn());
            this->menuStrip1->SuspendLayout();
            (cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->dataGridView1))->BeginInit();
            this->SuspendLayout();

            // menuStrip1
            this->menuStrip1->Items->AddRange(gcnew cli::array< System::Windows::Forms::ToolStripItem^  >(3) {
                this->ä³¿ToolStripMenuItem,
                    this->ïğîÏğîãğàìóToolStripMenuItem, this->âèõ³äToolStripMenuItem
            });
            this->menuStrip1->Location = System::Drawing::Point(0, 0);
            this->menuStrip1->Name = L"menuStrip1";
            this->menuStrip1->Size = System::Drawing::Size(1023, 28);
            this->menuStrip1->TabIndex = 0;
            this->menuStrip1->Text = L"menuStrip1";

            // ä³¿ToolStripMenuItem
            this->ä³¿ToolStripMenuItem->DropDownItems->AddRange(gcnew cli::array< System::Windows::Forms::ToolStripItem^  >(4) {
                this->çàâàíòàæèòèToolStripMenuItem,
                    this->îíToolStripMenuItem, this->äîäàòèToolStripMenuItem, this->âèäàëèòèToolStripMenuItem
            });
            this->ä³¿ToolStripMenuItem->Name = L"ä³¿ToolStripMenuItem";
            this->ä³¿ToolStripMenuItem->Size = System::Drawing::Size(41, 24);
            this->ä³¿ToolStripMenuItem->Text = L"Ä³¿";

            // çàâàíòàæèòèToolStripMenuItem
            this->çàâàíòàæèòèToolStripMenuItem->Name = L"çàâàíòàæèòèToolStripMenuItem";
            this->çàâàíòàæèòèToolStripMenuItem->Size = System::Drawing::Size(182, 26);
            this->çàâàíòàæèòèToolStripMenuItem->Text = L"Çàâàíòàæèòè";
            this->çàâàíòàæèòèToolStripMenuItem->Click += gcnew System::EventHandler(this, &MyForm::çàâàíòàæèòèToolStripMenuItem_Click);

            // îíToolStripMenuItem
            this->îíToolStripMenuItem->Name = L"îíToolStripMenuItem";
            this->îíToolStripMenuItem->Size = System::Drawing::Size(182, 26);
            this->îíToolStripMenuItem->Text = L"Îíîâèòè";
            this->îíToolStripMenuItem->Click += gcnew System::EventHandler(this, &MyForm::îíToolStripMenuItem_Click);

            // äîäàòèToolStripMenuItem
            this->äîäàòèToolStripMenuItem->Name = L"äîäàòèToolStripMenuItem";
            this->äîäàòèToolStripMenuItem->Size = System::Drawing::Size(182, 26);
            this->äîäàòèToolStripMenuItem->Text = L"Äîäàòè";
            this->äîäàòèToolStripMenuItem->Click += gcnew System::EventHandler(this, &MyForm::äîäàòèToolStripMenuItem_Click);

            // âèäàëèòèToolStripMenuItem
            this->âèäàëèòèToolStripMenuItem->Name = L"âèäàëèòèToolStripMenuItem";
            this->âèäàëèòèToolStripMenuItem->Size = System::Drawing::Size(182, 26);
            this->âèäàëèòèToolStripMenuItem->Text = L"Âèäàëèòè";
            this->âèäàëèòèToolStripMenuItem->Click += gcnew System::EventHandler(this, &MyForm::âèäàëèòèToolStripMenuItem_Click);

            // ïğîÏğîãğàìóToolStripMenuItem
            this->ïğîÏğîãğàìóToolStripMenuItem->Name = L"ïğîÏğîãğàìóToolStripMenuItem";
            this->ïğîÏğîãğàìóToolStripMenuItem->Size = System::Drawing::Size(124, 24);
            this->ïğîÏğîãğàìóToolStripMenuItem->Text = L"Ïğî ïğîãğàìó";
            this->ïğîÏğîãğàìóToolStripMenuItem->Click += gcnew System::EventHandler(this, &MyForm::ïğîÏğîãğàìóToolStripMenuItem_Click);

            // âèõ³äToolStripMenuItem
            this->âèõ³äToolStripMenuItem->Name = L"âèõ³äToolStripMenuItem";
            this->âèõ³äToolStripMenuItem->Size = System::Drawing::Size(60, 24);
            this->âèõ³äToolStripMenuItem->Text = L"Âèõ³ä";
            this->âèõ³äToolStripMenuItem->Click += gcnew System::EventHandler(this, &MyForm::âèõ³äToolStripMenuItem_Click);

            // dataGridView1
            this->dataGridView1->ColumnHeadersHeightSizeMode = System::Windows::Forms::DataGridViewColumnHeadersHeightSizeMode::AutoSize;
            this->dataGridView1->Columns->AddRange(gcnew cli::array< System::Windows::Forms::DataGridViewColumn^  >(5) {
                this->Id, this->ServiceName,
                    this->Price, this->Deadline, this->Detail
            });
            this->dataGridView1->Location = System::Drawing::Point(37, 71);
            this->dataGridView1->Name = L"dataGridView1";
            this->dataGridView1->RowHeadersWidth = 51;
            this->dataGridView1->RowTemplate->Height = 24;
            this->dataGridView1->Size = System::Drawing::Size(925, 180);
            this->dataGridView1->TabIndex = 1;

            // Id
            this->Id->HeaderText = L"Id";
            this->Id->MinimumWidth = 6;
            this->Id->Name = L"Id";
            this->Id->Width = 25;

            // ServiceName
            this->ServiceName->HeaderText = L"ServiceName";
            this->ServiceName->MinimumWidth = 6;
            this->ServiceName->Name = L"ServiceName";
            this->ServiceName->Width = 200;

            // Price
            this->Price->HeaderText = L"Price";
            this->Price->MinimumWidth = 6;
            this->Price->Name = L"Price";
            this->Price->Width = 125;

            // Deadline
            this->Deadline->HeaderText = L"Deadline";
            this->Deadline->MinimumWidth = 6;
            this->Deadline->Name = L"Deadline";
            this->Deadline->Width = 250;

            // Detail
            this->Detail->HeaderText = L"Detail";
            this->Detail->MinimumWidth = 6;
            this->Detail->Name = L"Detail";
            this->Detail->Width = 250;

            // MyForm
            this->AutoScaleDimensions = System::Drawing::SizeF(8, 16);
            this->AutoScaleMode = System::Windows::Forms::AutoScaleMode::Font;
            this->ClientSize = System::Drawing::Size(1023, 273);
            this->Controls->Add(this->dataGridView1);
            this->Controls->Add(this->menuStrip1);
            this->MainMenuStrip = this->menuStrip1;
            this->Name = L"MyForm";
            this->Text = L"Ïğèêëàä";
            this->Load += gcnew System::EventHandler(this, &MyForm::MyForm_Load);
            this->menuStrip1->ResumeLayout(false);
            this->menuStrip1->PerformLayout();
            (cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->dataGridView1))->EndInit();
            this->ResumeLayout(false);
            this->PerformLayout();
        }
#pragma endregion

    private: System::Void MyForm_Load(System::Object^ sender, System::EventArgs^ e) {
    }

    private: System::Void çàâàíòàæèòèToolStripMenuItem_Click(System::Object^ sender, System::EventArgs^ e) {
        try {
            String^ connString = L"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=db.accdb;";
            OleDbConnection^ dbConnection = gcnew OleDbConnection(connString);
            dbConnection->Open();

            String^ query = "SELECT * FROM [photo_table]";
            OleDbCommand^ dbCommand = gcnew OleDbCommand(query, dbConnection);
            OleDbDataReader^ dbReader = dbCommand->ExecuteReader();

            dataGridView1->Rows->Clear();

            while (dbReader->Read()) {
                dataGridView1->Rows->Add(dbReader["Id"], dbReader["ServiceName"], dbReader["Price"], dbReader["Deadline"], dbReader["Detail"]);
            }

            dbReader->Close();
            dbConnection->Close();
        }
        catch (Exception^ ex) {
            MessageBox::Show(ex->Message, "Error");
        }
    }

 private: System::Void äîäàòèToolStripMenuItem_Click(System::Object^ sender, System::EventArgs^ e) {
     try {
         // Äîäàºìî íîâèé ğÿäîê ó DataGridView
         int index = dataGridView1->Rows->Count - 1;

         // Ïåğåâ³ğêà, ÷è âñ³ íåîáõ³äí³ ïîëÿ çàïîâíåí³
         bool isDataMissing = false;

         for (int i = 1; i <= 4; i++) {
             if (dataGridView1->Rows[index]->Cells[i]->Value == nullptr ||
                 String::IsNullOrWhiteSpace(dataGridView1->Rows[index]->Cells[i]->Value->ToString())) {
                 isDataMissing = true;
                 break;
             }
         }

         if (isDataMissing) {
             MessageBox::Show(L"Íå âñ³ äàí³ º", L"Çâåğí³òü óâàãó");
             return;
         }

         // Ïåğåâ³ğêà, ùî Price º ÷èñëîì
         double price;
         if (!Double::TryParse(dataGridView1->Rows[index]->Cells[2]->Value->ToString(), price)) {
             MessageBox::Show(L"Ö³íà ìàº áóòè ÷èñëîì", L"Ïîìèëêà");
             return;
         }

         // Ï³äêëş÷åííÿ äî áàçè äàíèõ
         String^ connString = L"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=db.accdb;";
         OleDbConnection^ dbConnection = gcnew OleDbConnection(connString);
         dbConnection->Open();

         // Ôîğìóºìî çàïèò äëÿ âñòàâêè äàíèõ ó áàçó
         String^ query = "INSERT INTO [photo_table] (ServiceName, Price, Deadline, Detail) VALUES ('" +
             dataGridView1->Rows[index]->Cells[1]->Value->ToString() + "', '" +
             dataGridView1->Rows[index]->Cells[2]->Value->ToString() + "', '" +
             dataGridView1->Rows[index]->Cells[3]->Value->ToString() + "', '" +
             dataGridView1->Rows[index]->Cells[4]->Value->ToString() + "')";

         // Âèêîíóºìî çàïèò
         OleDbCommand^ dbCommand = gcnew OleDbCommand(query, dbConnection);
         dbCommand->ExecuteNonQuery();

         // Çàêğèâàºìî ç'ºäíàííÿ
         dbConnection->Close();

         MessageBox::Show(L"Äàí³ óñï³øíî äîäàí³", L"Óñï³õ");
     }
     catch (Exception^ ex) {
         MessageBox::Show(ex->Message, "Error");
     }
 }

    private: System::Void âèäàëèòèToolStripMenuItem_Click(System::Object^ sender, System::EventArgs^ e) {
        try {
            int index = dataGridView1->CurrentCell->RowIndex;

            String^ connString = L"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=db.accdb;";
            OleDbConnection^ dbConnection = gcnew OleDbConnection(connString);
            dbConnection->Open();

            String^ query = "DELETE FROM [photo_table] WHERE Id = " + dataGridView1->Rows[index]->Cells[0]->Value->ToString();
            OleDbCommand^ dbCommand = gcnew OleDbCommand(query, dbConnection);
            dbCommand->ExecuteNonQuery();

            dbConnection->Close();
            dataGridView1->Rows->RemoveAt(index);
        }
        catch (Exception^ ex) {
            MessageBox::Show(ex->Message, "Error");
        }
    }

    private: System::Void îíToolStripMenuItem_Click(System::Object^ sender, System::EventArgs^ e) {
        çàâàíòàæèòèToolStripMenuItem_Click(sender, e);
    }
           
    private: System::Void ïğîÏğîãğàìóToolStripMenuItem_Click(System::Object^ sender, System::EventArgs^ e) {
        MessageBox::Show(L"Öå ïğèêëàä ïğîãğàìè äëÿ ğîáîòè ç áàçîş äàíèõ ôîòî öåíòğó.", L"Ïğî ïğîãğàìó");
    }

    private: System::Void âèõ³äToolStripMenuItem_Click(System::Object^ sender, System::EventArgs^ e) {
        Application::Exit();
    }
    };
}