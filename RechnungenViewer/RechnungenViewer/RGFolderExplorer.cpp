#include "RGFolderExplorer.h"


void RGFolderExplorer::Update()
{
	ReadFolder();
}

void RGFolderExplorer::ReadFolder()
{
	auto di = gcnew System::IO::DirectoryInfo("\\\\gh-dc\\Buchhaltung\\Scanns\\Rechnungen");
	auto fileList = di->GetFiles("*.pdf", System::IO::SearchOption::AllDirectories);

	for each(auto file in fileList)
	{
		auto lvItem = gcnew System::Windows::Forms::ListViewItem(file->Name);
		lvItem->Tag = file;

		itemCollection.Add(lvItem);
	}
}
