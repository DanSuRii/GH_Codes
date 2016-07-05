#pragma once
ref class RGFolderExplorer
{
	System::Collections::Generic::List<System::Windows::Forms::ListViewItem^> itemCollection;
public:
	static property RGFolderExplorer^ Instance
	{
		RGFolderExplorer^ get() { return %_Instance; }
	}

	property decltype(itemCollection) ^ Items
	{
		decltype(itemCollection) ^ get() { return %itemCollection; }
	}

	//virtual ~RGFolderExplorer();
	
	void Update();

private:
	RGFolderExplorer()
		: itemCollection(gcnew System::Collections::Generic::List<System::Windows::Forms::ListViewItem^>())
	{
		ReadFolder();
	};

	void ReadFolder();
	static RGFolderExplorer _Instance;	
};

