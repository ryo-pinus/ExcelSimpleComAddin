

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


 /* File created by MIDL compiler version 8.00.0603 */
/* at Sun Jun 14 22:24:39 2015
 */
/* Compiler settings for ExcelSimpleComAddin.idl:
    Oicf, W1, Zp8, env=Win64 (32b run), target_arch=AMD64 8.00.0603 
    protocol : dce , ms_ext, c_ext, robust
    error checks: allocation ref bounds_check enum stub_data 
    VC __declspec() decoration level: 
         __declspec(uuid()), __declspec(selectany), __declspec(novtable)
         DECLSPEC_UUID(), MIDL_INTERFACE()
*/
/* @@MIDL_FILE_HEADING(  ) */

#pragma warning( disable: 4049 )  /* more than 64k source lines */


/* verify that the <rpcndr.h> version is high enough to compile this file*/
#ifndef __REQUIRED_RPCNDR_H_VERSION__
#define __REQUIRED_RPCNDR_H_VERSION__ 475
#endif

#include "rpc.h"
#include "rpcndr.h"

#ifndef __RPCNDR_H_VERSION__
#error this stub requires an updated version of <rpcndr.h>
#endif // __RPCNDR_H_VERSION__

#ifndef COM_NO_WINDOWS_H
#include "windows.h"
#include "ole2.h"
#endif /*COM_NO_WINDOWS_H*/

#ifndef __ExcelSimpleComAddin_i_h__
#define __ExcelSimpleComAddin_i_h__

#if defined(_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

/* Forward Declarations */ 

#ifndef __IMyConnect_FWD_DEFINED__
#define __IMyConnect_FWD_DEFINED__
typedef interface IMyConnect IMyConnect;

#endif 	/* __IMyConnect_FWD_DEFINED__ */


#ifndef __MyConnect_FWD_DEFINED__
#define __MyConnect_FWD_DEFINED__

#ifdef __cplusplus
typedef class MyConnect MyConnect;
#else
typedef struct MyConnect MyConnect;
#endif /* __cplusplus */

#endif 	/* __MyConnect_FWD_DEFINED__ */


/* header files for imported files */
#include "oaidl.h"
#include "ocidl.h"

#ifdef __cplusplus
extern "C"{
#endif 


#ifndef __IMyConnect_INTERFACE_DEFINED__
#define __IMyConnect_INTERFACE_DEFINED__

/* interface IMyConnect */
/* [unique][nonextensible][oleautomation][uuid][object] */ 


EXTERN_C const IID IID_IMyConnect;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("1DB68E7D-D1D6-420E-9638-F9FD64FF3B8A")
    IMyConnect : public IUnknown
    {
    public:
    };
    
    
#else 	/* C style interface */

    typedef struct IMyConnectVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IMyConnect * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IMyConnect * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IMyConnect * This);
        
        END_INTERFACE
    } IMyConnectVtbl;

    interface IMyConnect
    {
        CONST_VTBL struct IMyConnectVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IMyConnect_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IMyConnect_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IMyConnect_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IMyConnect_INTERFACE_DEFINED__ */



#ifndef __ExcelSimpleComAddinLib_LIBRARY_DEFINED__
#define __ExcelSimpleComAddinLib_LIBRARY_DEFINED__

/* library ExcelSimpleComAddinLib */
/* [version][uuid] */ 


EXTERN_C const IID LIBID_ExcelSimpleComAddinLib;

EXTERN_C const CLSID CLSID_MyConnect;

#ifdef __cplusplus

class DECLSPEC_UUID("693DC738-F76F-451B-AE5B-75388ACD5EF0")
MyConnect;
#endif
#endif /* __ExcelSimpleComAddinLib_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif


