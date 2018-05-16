#include "RGFolderExplorer.h"


void RGFolderExplorer::Update()
{
	ReadFolder();
}

void RGFolderExplorer::ReadFolder()
{
	//optimizierung

	//check the time begin
	auto di = gcnew System::IO::DirectoryInfo("\\\\gh-dc\\Buchhaltung\\Scanns\\Rechnungen");
	//check end 1st, check 2nd begin
	auto fileList = di->GetFiles("*.pdf", System::IO::SearchOption::AllDirectories);
	//check 2nd end

	//check 3rd begin
	for each(auto file in fileList)
	{
		auto lvItem = gcnew System::Windows::Forms::ListViewItem(file->Name);
		lvItem->Tag = file;

		itemCollection.Add(lvItem);
	}
	//check 3rd end
}
