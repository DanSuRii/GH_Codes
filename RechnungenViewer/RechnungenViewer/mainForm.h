#pragma once

namespace RechnungenViewer {

	using namespace System;
	using namespace System::ComponentModel;
	using namespace System::Collections;
	using namespace System::Windows::Forms;
	using namespace System::Data;
	using namespace System::Drawing;

	/// <summary>
	/// Zusammenfassung für mainForm
	/// </summary>
	public ref class mainForm : public System::Windows::Forms::Form
	{
	public:
		mainForm(void)
		{
			InitializeComponent();
			//
			//TODO: Konstruktorcode hier hinzufügen.
			//
		}

	protected:
		/// <summary>
		/// Verwendete Ressourcen bereinigen.
		/// </summary>
		~mainForm()
		{
			if (components)
			{
				delete components;
			}
		}
	private: System::IO::FileSystemWatcher^  fileSystemWatcher1;
	protected:
	private: System::Windows::Forms::ListView^  FileList;
	private: System::Windows::Forms::ColumnHeader^  FileName;
	private: System::Windows::Forms::ColumnHeader^  Folder;
	private: System::Windows::Forms::TextBox^  SearchText;


	private: System::Windows::Forms::SaveFileDialog^  saveFileDialog1;
	private: System::Windows::Forms::Label^  label1;
	private: System::Windows::Forms::PrintPreviewDialog^  printPreviewDialog1;

	private:
		/// <summary>
		/// Erforderliche Designervariable.
		/// </summary>
		System::ComponentModel::Container ^components;

#pragma region Windows Form Designer generated code
		/// <summary>
		/// Erforderliche Methode für die Designerunterstützung.
		/// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
		/// </summary>
		void InitializeComponent(void)
		{
			System::ComponentModel::ComponentResourceManager^  resources = (gcnew System::ComponentModel::ComponentResourceManager(mainForm::typeid));
			this->fileSystemWatcher1 = (gcnew System::IO::FileSystemWatcher());
			this->saveFileDialog1 = (gcnew System::Windows::Forms::SaveFileDialog());
			this->SearchText = (gcnew System::Windows::Forms::TextBox());
			this->FileList = (gcnew System::Windows::Forms::ListView());
			this->FileName = (gcnew System::Windows::Forms::ColumnHeader());
			this->Folder = (gcnew System::Windows::Forms::ColumnHeader());
			this->label1 = (gcnew System::Windows::Forms::Label());
			this->printPreviewDialog1 = (gcnew System::Windows::Forms::PrintPreviewDialog());
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->fileSystemWatcher1))->BeginInit();
			this->SuspendLayout();
			// 
			// fileSystemWatcher1
			// 
			this->fileSystemWatcher1->EnableRaisingEvents = true;
			this->fileSystemWatcher1->SynchronizingObject = this;
			// 
			// SearchText
			// 
			this->SearchText->Location = System::Drawing::Point(163, 42);
			this->SearchText->Name = L"SearchText";
			this->SearchText->Size = System::Drawing::Size(160, 22);
			this->SearchText->TabIndex = 1;
			this->SearchText->TextChanged += gcnew System::EventHandler(this, &mainForm::SearchText_TextChanged);
			// 
			// FileList
			// 
			this->FileList->Columns->AddRange(gcnew cli::array< System::Windows::Forms::ColumnHeader^  >(2) { this->FileName, this->Folder });
			this->FileList->GridLines = true;
			this->FileList->Location = System::Drawing::Point(57, 162);
			this->FileList->MultiSelect = false;
			this->FileList->Name = L"FileList";
			this->FileList->Size = System::Drawing::Size(266, 238);
			this->FileList->TabIndex = 2;
			this->FileList->UseCompatibleStateImageBehavior = false;
			this->FileList->View = System::Windows::Forms::View::Details;
			this->FileList->ItemSelectionChanged += gcnew System::Windows::Forms::ListViewItemSelectionChangedEventHandler(this, &mainForm::FileList_ItemSelectionChanged);
			this->FileList->DoubleClick += gcnew System::EventHandler(this, &mainForm::FileList_DoubleClick);
			// 
			// FileName
			// 
			this->FileName->Text = L"FileName";
			this->FileName->Width = 90;
			// 
			// Folder
			// 
			this->Folder->Text = L"Folder";
			this->Folder->Width = 90;
			// 
			// label1
			// 
			this->label1->AutoSize = true;
			this->label1->Location = System::Drawing::Point(57, 42);
			this->label1->Name = L"label1";
			this->label1->Size = System::Drawing::Size(56, 17);
			this->label1->TabIndex = 3;
			this->label1->Text = L"Suchen";
			// 
			// printPreviewDialog1
			// 
			this->printPreviewDialog1->AutoScrollMargin = System::Drawing::Size(0, 0);
			this->printPreviewDialog1->AutoScrollMinSize = System::Drawing::Size(0, 0);
			this->printPreviewDialog1->ClientSize = System::Drawing::Size(400, 300);
			this->printPreviewDialog1->Enabled = true;
			this->printPreviewDialog1->Icon = (cli::safe_cast<System::Drawing::Icon^>(resources->GetObject(L"printPreviewDialog1.Icon")));
			this->printPreviewDialog1->Name = L"printPreviewDialog1";
			this->printPreviewDialog1->Visible = false;
			// 
			// mainForm
			// 
			this->AutoScaleDimensions = System::Drawing::SizeF(8, 16);
			this->AutoScaleMode = System::Windows::Forms::AutoScaleMode::Font;
			this->ClientSize = System::Drawing::Size(404, 440);
			this->Controls->Add(this->label1);
			this->Controls->Add(this->FileList);
			this->Controls->Add(this->SearchText);
			this->FormBorderStyle = System::Windows::Forms::FormBorderStyle::FixedSingle;
			this->MaximizeBox = false;
			this->Name = L"mainForm";
			this->ShowIcon = false;
			this->SizeGripStyle = System::Windows::Forms::SizeGripStyle::Hide;
			this->Text = L"RechnungenViewer";
			this->Shown += gcnew System::EventHandler(this, &mainForm::mainForm_Shown);
			this->SizeChanged += gcnew System::EventHandler(this, &mainForm::mainForm_SizeChanged);
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->fileSystemWatcher1))->EndInit();
			this->ResumeLayout(false);
			this->PerformLayout();

		}
#pragma endregion
private: System::Void mainForm_Shown(System::Object^  sender, System::EventArgs^  e);
private: System::Void mainForm_SizeChanged(System::Object^  sender, System::EventArgs^  e);

		 // Read folder and listing Files
		 //void ReadFolder();
private: System::Void FileList_ItemSelectionChanged(System::Object^  sender, System::Windows::Forms::ListViewItemSelectionChangedEventArgs^  e);
private: System::Void FileList_DoubleClick(System::Object^  sender, System::EventArgs^  e);
private: System::Void SearchText_TextChanged(System::Object^  sender, System::EventArgs^  e);
};
}
