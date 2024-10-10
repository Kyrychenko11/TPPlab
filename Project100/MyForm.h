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
    private: System::Windows::Forms::ToolStripMenuItem^ 䳿ToolStripMenuItem;
    private: System::Windows::Forms::ToolStripMenuItem^ �����������ToolStripMenuItem;
    private: System::Windows::Forms::ToolStripMenuItem^ �����������ToolStripMenuItem;
    private: System::Windows::Forms::ToolStripMenuItem^ �����ToolStripMenuItem;
    private: System::Windows::Forms::DataGridView^ dataGridView1;
    private: System::Windows::Forms::ToolStripMenuItem^ ��ToolStripMenuItem;
    private: System::Windows::Forms::ToolStripMenuItem^ ������ToolStripMenuItem;
    private: System::Windows::Forms::ToolStripMenuItem^ ��������ToolStripMenuItem;
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
            this->䳿ToolStripMenuItem = (gcnew System::Windows::Forms::ToolStripMenuItem());
            this->�����������ToolStripMenuItem = (gcnew System::Windows::Forms::ToolStripMenuItem());
            this->��ToolStripMenuItem = (gcnew System::Windows::Forms::ToolStripMenuItem());
            this->������ToolStripMenuItem = (gcnew System::Windows::Forms::ToolStripMenuItem());
            this->��������ToolStripMenuItem = (gcnew System::Windows::Forms::ToolStripMenuItem());
            this->�����������ToolStripMenuItem = (gcnew System::Windows::Forms::ToolStripMenuItem());
            this->�����ToolStripMenuItem = (gcnew System::Windows::Forms::ToolStripMenuItem());
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
                this->䳿ToolStripMenuItem,
                    this->�����������ToolStripMenuItem, this->�����ToolStripMenuItem
            });
            this->menuStrip1->Location = System::Drawing::Point(0, 0);
            this->menuStrip1->Name = L"menuStrip1";
            this->menuStrip1->Size = System::Drawing::Size(1023, 28);
            this->menuStrip1->TabIndex = 0;
            this->menuStrip1->Text = L"menuStrip1";

            // 䳿ToolStripMenuItem
            this->䳿ToolStripMenuItem->DropDownItems->AddRange(gcnew cli::array< System::Windows::Forms::ToolStripItem^  >(4) {
                this->�����������ToolStripMenuItem,
                    this->��ToolStripMenuItem, this->������ToolStripMenuItem, this->��������ToolStripMenuItem
            });
            this->䳿ToolStripMenuItem->Name = L"䳿ToolStripMenuItem";
            this->䳿ToolStripMenuItem->Size = System::Drawing::Size(41, 24);
            this->䳿ToolStripMenuItem->Text = L"ĳ�";

            // �����������ToolStripMenuItem
            this->�����������ToolStripMenuItem->Name = L"�����������ToolStripMenuItem";
            this->�����������ToolStripMenuItem->Size = System::Drawing::Size(182, 26);
            this->�����������ToolStripMenuItem->Text = L"�����������";
            this->�����������ToolStripMenuItem->Click += gcnew System::EventHandler(this, &MyForm::�����������ToolStripMenuItem_Click);

            // ��ToolStripMenuItem
            this->��ToolStripMenuItem->Name = L"��ToolStripMenuItem";
            this->��ToolStripMenuItem->Size = System::Drawing::Size(182, 26);
            this->��ToolStripMenuItem->Text = L"�������";
            this->��ToolStripMenuItem->Click += gcnew System::EventHandler(this, &MyForm::��ToolStripMenuItem_Click);

            // ������ToolStripMenuItem
            this->������ToolStripMenuItem->Name = L"������ToolStripMenuItem";
            this->������ToolStripMenuItem->Size = System::Drawing::Size(182, 26);
            this->������ToolStripMenuItem->Text = L"������";
            this->������ToolStripMenuItem->Click += gcnew System::EventHandler(this, &MyForm::������ToolStripMenuItem_Click);

            // ��������ToolStripMenuItem
            this->��������ToolStripMenuItem->Name = L"��������ToolStripMenuItem";
            this->��������ToolStripMenuItem->Size = System::Drawing::Size(182, 26);
            this->��������ToolStripMenuItem->Text = L"��������";
            this->��������ToolStripMenuItem->Click += gcnew System::EventHandler(this, &MyForm::��������ToolStripMenuItem_Click);

            // �����������ToolStripMenuItem
            this->�����������ToolStripMenuItem->Name = L"�����������ToolStripMenuItem";
            this->�����������ToolStripMenuItem->Size = System::Drawing::Size(124, 24);
            this->�����������ToolStripMenuItem->Text = L"��� ��������";
            this->�����������ToolStripMenuItem->Click += gcnew System::EventHandler(this, &MyForm::�����������ToolStripMenuItem_Click);

            // �����ToolStripMenuItem
            this->�����ToolStripMenuItem->Name = L"�����ToolStripMenuItem";
            this->�����ToolStripMenuItem->Size = System::Drawing::Size(60, 24);
            this->�����ToolStripMenuItem->Text = L"�����";
            this->�����ToolStripMenuItem->Click += gcnew System::EventHandler(this, &MyForm::�����ToolStripMenuItem_Click);

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
            this->Text = L"�������";
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

    private: System::Void �����������ToolStripMenuItem_Click(System::Object^ sender, System::EventArgs^ e) {
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

 private: System::Void ������ToolStripMenuItem_Click(System::Object^ sender, System::EventArgs^ e) {
     try {
         // ������ ����� ����� � DataGridView
         int index = dataGridView1->Rows->Count - 1;

         // ��������, �� �� �������� ���� ��������
         bool isDataMissing = false;

         for (int i = 1; i <= 4; i++) {
             if (dataGridView1->Rows[index]->Cells[i]->Value == nullptr ||
                 String::IsNullOrWhiteSpace(dataGridView1->Rows[index]->Cells[i]->Value->ToString())) {
                 isDataMissing = true;
                 break;
             }
         }

         if (isDataMissing) {
             MessageBox::Show(L"�� �� ��� �", L"������� �����");
             return;
         }

         // ��������, �� Price � ������
         double price;
         if (!Double::TryParse(dataGridView1->Rows[index]->Cells[2]->Value->ToString(), price)) {
             MessageBox::Show(L"ֳ�� �� ���� ������", L"�������");
             return;
         }

         // ϳ��������� �� ���� �����
         String^ connString = L"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=db.accdb;";
         OleDbConnection^ dbConnection = gcnew OleDbConnection(connString);
         dbConnection->Open();

         // ������� ����� ��� ������� ����� � ����
         String^ query = "INSERT INTO [photo_table] (ServiceName, Price, Deadline, Detail) VALUES ('" +
             dataGridView1->Rows[index]->Cells[1]->Value->ToString() + "', '" +
             dataGridView1->Rows[index]->Cells[2]->Value->ToString() + "', '" +
             dataGridView1->Rows[index]->Cells[3]->Value->ToString() + "', '" +
             dataGridView1->Rows[index]->Cells[4]->Value->ToString() + "')";

         // �������� �����
         OleDbCommand^ dbCommand = gcnew OleDbCommand(query, dbConnection);
         dbCommand->ExecuteNonQuery();

         // ��������� �'�������
         dbConnection->Close();

         MessageBox::Show(L"��� ������ �����", L"����");
     }
     catch (Exception^ ex) {
         MessageBox::Show(ex->Message, "Error");
     }
 }

    private: System::Void ��������ToolStripMenuItem_Click(System::Object^ sender, System::EventArgs^ e) {
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

    private: System::Void ��ToolStripMenuItem_Click(System::Object^ sender, System::EventArgs^ e) {
        �����������ToolStripMenuItem_Click(sender, e);
    }
           
    private: System::Void �����������ToolStripMenuItem_Click(System::Object^ sender, System::EventArgs^ e) {
        MessageBox::Show(L"�� ������� �������� ��� ������ � ����� ����� ���� ������.", L"��� ��������");
    }

    private: System::Void �����ToolStripMenuItem_Click(System::Object^ sender, System::EventArgs^ e) {
        Application::Exit();
    }
    };
}