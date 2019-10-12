/*
com-COUNTER for OpenVBS CC BY @taisukef
base src by OpenVBS SCRIPTING.cpp CC BY NaturalStyle PREMIUM


GUID
https://www.guidgenerator.com/online-guid-generator.aspx
generated Uppercase
92DCA431-FA00-414D-82C0-76D5313BF5C4

g++ -Wall -I../src -std=c++14 -O0 -shared COUNTER.cpp ../src/npole.cpp -o COUNTER.so
echo {92DCA431-FA00-414D-82C0-76D5313BF5C4} > COUNTER
ln -s COUNTER.so {92DCA431-FA00-414D-82C0-76D5313BF5C4}

*/

#include <npole.h>
#include <string>
#include <iostream>
#include <fstream>

#define FILE "counter.txt"

class Counter : public IDispatch{
private:
    ULONG       m_refc   = 1;

public:
    int n;

public:
    Counter() {
        /*wprintf(L"%s\n", __func__);*/
        n = 0;
        load();
    }
    virtual ~Counter() {
        /*wprintf(L"%s\n", __func__);*/
    }
protected:
    void load() {
        std::ifstream fin;
        fin.open(FILE);
        if (!fin.fail()) {
            std::string str;
            fin >> str;
            n = atoi(str.c_str());
        }
        fin.close();
    }
    void save() {
        std::ofstream fout;
        fout.open(FILE);
        fout << n;
        fout.close();
    }

public:
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject){ return E_NOTIMPL; }
    ULONG STDMETHODCALLTYPE AddRef(){ return ++m_refc; }
    ULONG STDMETHODCALLTYPE Release(){ if(!--m_refc){ delete this; return 0; } return m_refc; }

    HRESULT STDMETHODCALLTYPE GetTypeInfoCount(UINT *pctinfo){ return E_NOTIMPL; }
    HRESULT STDMETHODCALLTYPE GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo){ return E_NOTIMPL; }
    HRESULT STDMETHODCALLTYPE GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId){
        if (_wcsicmp(*rgszNames, L"Inc") == 0){
            *rgDispId = 0;
        } else if(_wcsicmp(*rgszNames, L"Dec") == 0){
            *rgDispId = 1;
        } else if(_wcsicmp(*rgszNames, L"Get") == 0){
            *rgDispId = 2;
        } else if(_wcsicmp(*rgszNames, L"Reset") == 0){
            *rgDispId = 3;
        } else {
            wprintf(L"###%s: Implement here '%s' line %d. (%ls)\n", __func__, __FILE__, __LINE__, *rgszNames);
            return DISP_E_MEMBERNOTFOUND;
        }
        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) {
        if (dispIdMember == 0) {
            n++;
            save();
        } else if(dispIdMember == 1) {
            n--;
            save();
        } else if (dispIdMember == 2){
            pVarResult->vt = VT_I8;
            pVarResult->llVal = n;
        } else if (dispIdMember == 3){
            n = 0;
            save();
        } else {
            return DISP_E_MEMBERNOTFOUND;
        }
        return S_OK;
    }
};

class CFactory : public IClassFactory{
private:
    ULONG       m_refc   = 1;

public:
    CFactory(){}
    virtual ~CFactory(){}

public:
    HRESULT QueryInterface(REFIID riid, void** ppvObject){ return E_NOTIMPL; }
    ULONG AddRef(){ return ++m_refc; }
    ULONG Release(){ if(!--m_refc){ delete this; return 0; } return m_refc; }

    HRESULT CreateInstance(IUnknown *pUnkOuter, REFIID riid, void **ppvObject){
        *ppvObject = new Counter;
        return S_OK;
    }    
    HRESULT LockServer(BOOL fLock){
        m_refc += fLock ? 1 : -1;
        return S_OK;
    }
} g_oFactory;



extern "C"
HRESULT DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID *ppv){
    g_oFactory.AddRef();
    *ppv = (IClassFactory*)&g_oFactory;

    return S_OK;
}

extern "C"
HRESULT DllCanUnloadNow(){
    return S_FALSE;
}
