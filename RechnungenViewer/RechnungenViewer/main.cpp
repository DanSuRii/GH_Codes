#if !defined(_DEBUG)
#include <windows.h>
#endif


using namespace System;
using namespace System::Threading::Tasks;



#include <iostream>
#include "mainForm.h"




class A
{
public:
	void close() { std::cout << "dispose class A" << std::endl; }
};

class B
{
public:
	void quit() { std::cout << "dispose class B" << std::endl; }
};

template< class T >
class RAII
{
	T& t;
public:
	RAII(T& rhs):t(rhs) {};

	virtual ~RAII()	
	{ 
		Dispose();
	}

	void Dispose();
};

int main(array<System::String ^> ^args)
{
	//System::Console::WriteLine("Hello world");
#ifdef _DEBUG
	A a;
	B b;

	A* ap;
	{
		RAII<A> raii(a);
		RAII<B> raii2(b);

		//RAII<A*> raii(ap);
	
	}
#endif

	//system("pause");

	//System::Console::SetWindowSize(0, 0);

#if !defined(_DEBUG)
	FreeConsole();
#endif

	RechnungenViewer::mainForm Form;
	Form.ShowDialog();
	return 0;
}

template<> void RAII<A>::Dispose()
{
	t.close();
}

template<> void RAII<B>::Dispose()
{
	t.quit();
}