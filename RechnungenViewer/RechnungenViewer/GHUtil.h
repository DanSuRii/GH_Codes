#pragma once

namespace GH_UTIL
{
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

	ref class SUtil
	{
	public:
		static property SUtil^ Instance
		{
			SUtil^ get() { return %_Instance; }
		}

		System::Predicate<System::Windows::Forms::ListViewItem^>^ LvItemCompare;
	private:
		SUtil()
			//: LvItemCompare(gcnew System::Predicate<System::Windows::Forms::ListViewItem^>(gcnew CCompLvItem(searchText), &CCompLvItem::Comp));
		{
		}

		static SUtil		_Instance;
	};
}