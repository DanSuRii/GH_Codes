#include "mainForm.h"
#include "RAII.h"
#include "RGFolderExplorer.h"
#include "GHUtil.h"



inline System::Void RechnungenViewer::mainForm::mainForm_Shown(System::Object ^ sender, System::EventArgs ^ e)
{	
#ifdef _DEBUG
	printPreviewDialog1->ControlBox = false;
	printPreviewDialog1->Show();
#endif

	RGFolderExplorer^ rgInit = RGFolderExplorer::Instance;
	FileList->Items->AddRange( rgInit->Items->ToArray() );

}

inline System::Void RechnungenViewer::mainForm::mainForm_SizeChanged(System::Object ^ sender, System::EventArgs ^ e)
{
	printPreviewDialog1->WindowState = this->WindowState;
}


// Unused, code moved 20160704 RGFolderExplorer
// Read folder and listing Files
#if false
inline void RechnungenViewer::mainForm::ReadFolder()
{
	//;
	//auto strFolder = gcnew System::String("C:\\Users\\A.Roennburg\\Documents\\SGS\\Report\\2016");
	//auto di = gcnew System::IO::DirectoryInfo("C:\\Users\\A.Roennburg\\Documents\\SGS\\Report\\2016");

	auto strFolder = gcnew System::String("\\\\gh-dc\\Buchhaltung\\Scanns\\Rechnungen");
	auto di = gcnew System::IO::DirectoryInfo("\\\\gh-dc\\Buchhaltung\\Scanns\\Rechnungen");

	//O:\Scanns\Rechnungen
	auto fileList = di->GetFiles("*.pdf", System::IO::SearchOption::AllDirectories);
	//auto itemCollection = gcnew System::Windows::Forms::ListView::ListViewItemCollection();
	//auto itemCollection = gcnew System::Collections::Generic::LinkedList<System::Windows::Forms::ListViewItem^>( );
	//auto itemCollection = gcnew Array<System::Windows::Forms::ListViewItem^>( );
	auto itemCollection = gcnew System::Collections::Generic::List<System::Windows::Forms::ListViewItem^>();


	for each(auto file in fileList)
	{
		auto lvItem = gcnew ListViewItem(file->Name);
		lvItem->Tag = file;

		itemCollection->Add(lvItem);
	}

	FileList->Items->AddRange(itemCollection->ToArray());
}

#endif // false
inline System::Void RechnungenViewer::mainForm::FileList_ItemSelectionChanged(System::Object ^ sender, System::Windows::Forms::ListViewItemSelectionChangedEventArgs ^ e)
{
	//TODO: Made it preview
	/*
	auto itemSelected = FileList->SelectedItems[0];
	auto fileInfo = (System::IO::FileInfo^)(itemSelected->Tag);
	auto strPath = fileInfo->FullName;
	//printPreviewDialog1->Document->PrintPage += gcnew Drawing::Printing::PrintPageEventHandler();
	auto streamToPrint = gcnew System::IO::StreamReader( strPath );
	GH_UTIL::DiposeToClose< decltype(streamToPrint) > raii(streamToPrint);

	System::Drawing::Printing::PrintDocument;
	
	auto ExcelApp = gcnew Microsoft::Office::Interop::Excel::Application();
	GH_UTIL::RAII< decltype(ExcelApp) > raiiExcelApp(ExcelApp);
	*/
}

inline System::Void RechnungenViewer::mainForm::FileList_DoubleClick(System::Object ^ sender, System::EventArgs ^ e)
{
	System::IO::FileInfo^ file = (System::IO::FileInfo^)FileList->SelectedItems[0]->Tag;

	System::Diagnostics::Process::Start(file->FullName);
}

public delegate bool CompareLvItem(System::Windows::Forms::ListViewItem^, System::String^);
bool FnCompareLvItem(System::Windows::Forms::ListViewItem^ lvItem, System::String^ strComp)
{
	return System::String::Compare(lvItem->Name, strComp);
}

ref class CCompLvItem
{
	System::String^ strToCmp;
public:

	CCompLvItem(System::String^ rhs) :strToCmp(rhs) {}
	bool Comp(System::Windows::Forms::ListViewItem^ lvItem)
	{
		//return System::String::Compare(lvItem->Name, strToCmp);
		return lvItem->Text->StartsWith(strToCmp);
	}
};

inline System::Void RechnungenViewer::mainForm::SearchText_TextChanged(System::Object ^ sender, System::EventArgs ^ e)
{
	auto searchText = SearchText->Text;
	RGFolderExplorer^ RGExp = RGFolderExplorer::Instance;

	//FileList->BeginUpdate();
	{
		//GH_UTIL::DisposeToUpdate<> updater(FileList);
		FileList->Items->Clear();
		if (0 == searchText->Length)
		{
			FileList->Items->AddRange(RGExp->Items->ToArray());
		}
		else
		{
			//c++cli does not support Lambda
			//auto func = gcnew CompareLvItem(FnCompareLvItem);
			//auto expFind = [ searchText ](ListViewItem^ item) { item };

			FileList->Items->AddRange (
				RGExp->Items->FindAll(
					gcnew Predicate<ListViewItem^>( gcnew CCompLvItem(searchText), &CCompLvItem::Comp ))->ToArray()
			);

			//FileList->Items->
		}
	}
}

//Old SearchText_TextChanged
#if 0
inline System::Void RechnungenViewer::mainForm::SearchText_TextChanged(System::Object ^ sender, System::EventArgs ^ e)
{
	auto searchText = SearchText->Text;

	FileList->BeginUpdate();
	FileList->Items->Clear();

	auto di = gcnew System::IO::DirectoryInfo("\\\\gh-dc\\Buchhaltung\\Scanns\\Rechnungen");

	//O:\Scanns\Rechnungen
	auto fileList = di->GetFiles(searchText + "*.pdf", System::IO::SearchOption::AllDirectories);
	//auto itemCollection = gcnew System::Windows::Forms::ListView::ListViewItemCollection();
	//auto itemCollection = gcnew System::Collections::Generic::LinkedList<System::Windows::Forms::ListViewItem^>( );
	//auto itemCollection = gcnew Array<System::Windows::Forms::ListViewItem^>( );
	auto itemCollection = gcnew System::Collections::Generic::List<System::Windows::Forms::ListViewItem^>();


	for each(auto file in fileList)
	{
		auto lvItem = gcnew ListViewItem(file->Name);
		lvItem->Tag = file;

		itemCollection->Add(lvItem);
	}

	FileList->Items->AddRange(itemCollection->ToArray());

	FileList->EndUpdate();
	//FileList->Focus();
	//FileList->Items[0]->Selected = true;
}
#endif // Codes expired