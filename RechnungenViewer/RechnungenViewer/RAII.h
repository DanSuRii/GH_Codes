#pragma once

namespace GH_UTIL
{
	template<class T>
	class RAII
	{
	public:
		RAII(T& t) :obj(t) {}
		virtual ~RAII() {
			Destruct();
		}

		virtual void Dipose();

	protected:
		T& obj;
	private:
		RAII();
		//caused by Destructor cannot call dynamic binded function.
		void Destruct() { return Dipose(); }
	};	
	
	template<class T>
	void RAII<T>::Dipose()
	{
	}

	/*
	template<> void RAII<Microsoft::Office::Interop::Excel::ApplicationClass ^>::Dipose()
	{
		obj->Quit();
	}
	*/

	template<> void RAII<System::Windows::Forms::ListView ^>::Dipose()
	{
		obj->EndUpdate();
	}

	template<class T>
	class DiposeToClose : public RAII<T>
	{
	public:
		DiposeToClose(T& t) :RAII<T>(t) {}
		virtual ~DiposeToClose() { obj->Close(); }
		
		//never called by dynamic bindinng, with no caused :(	
		//virtual void Dipose() { obj->Close(); }
	};

	template<class T>
	class DiposeToQuit : public RAII<T>
	{
	public:
		DiposeToQuit(T& t) :RAII<T>(t) {}
		virtual ~DiposeToQuit() { obj->Quit(); }

		//never called by dynamic bindinng, with no caused :(	
		//virtual void Dipose() { obj->Close(); }
	};

	template<class T = System::Windows::Forms::ListView ^>
	class DisposeToUpdate : public RAII<T>
	{
	public:
		DisposeToUpdate(T t) : RAII<T>(t) {}
	};
}
