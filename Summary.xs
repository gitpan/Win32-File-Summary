#ifdef __cplusplus
extern "C" {
#endif
#include "windows.h"
#include "Objidl.h"
#include "defines.h"

//#undef Summary
#undef THIS 

#include "EXTERN.h"
#include "perl.h"
#include "XSUB.h"
#ifdef __cplusplus
}
#endif

/* Is an argument C<[]>, meaning we should pass C<NULL>? */
/*#define null_arg(sv)	(  SvROK(sv)  &&  SVt_PVAV == SvTYPE(SvRV(sv))	\
			   &&  -1 == av_len((AV*)SvRV(sv))  )
*/
class  
Summary {
	public: 
		Summary(char *File)
		{
			
			char tmp[4];
			m_IsOOo = 0;
			int len = 0;
			MultiByteToWideChar(CP_ACP, 0, File, -1, m_File, (sizeof(m_File)/sizeof(WCHAR)));
			m_hr = S_OK;
			m_ipStg = NULL;
			m_hv = (HV* )newHV();
			len = strlen(File);
			tmp[0] = File[len-3];
			tmp[1] = File[len-2];
			tmp[2] = File[len-1];
			tmp[3] = '\0';
			//printf("tmp : %s\n", tmp);
			m_IsOOo = 0;
			if(!strcmp(tmp,"sxw")) m_IsOOo = 1;
			if(!strcmp(tmp,"sxc")) m_IsOOo = 1;
			if(!strcmp(tmp,"ods")) m_IsOOo = 1;
		}
		~Summary();
		int IsWin2000OrNT(void);
		int IsStgFile(void);
		int IsNTFS(void);
		SV* Read(void);
		int Write(void);
		SV* GetError(void);
		int IsOOoFile(void) { return(m_IsOOo); }
	private:
		void HrToString(HRESULT hr, char *string);
		void PropertyPIDToCaption(PROPSPEC propspec, char *title);
		SV* ReadOOo(SV* ptResult);
		void SetErr(char *msg);
	private:
		wchar_t m_File[2048];
		char m_perror[1024];
		HRESULT m_hr;
		IPropertySetStorage *m_ipStg;
		HV* m_hv;
		int m_IsOOo;
		

};


void 
Summary::SetErr(char *msg)
{
	char tmp[1000] = { '\0' };
	// SecureZeroMemory(&tmp,tmplen);
	HrToString(m_hr, tmp);
	strcpy(m_perror, msg);
	strcat(m_perror, tmp);
	

}

SV*
Summary::GetError(void)
{
	//SV* err = NEWSV(0,0);
	//err = newRV_noinc(newSVpvn(m_perror, strlen(m_perror)));
	return (newRV_noinc(newSVpvn(m_perror, strlen(m_perror))));
	//return (err);
}

Summary::~Summary()
{ 
	if( m_ipStg ) m_ipStg->Release(); 
	printf("Closing\n");
}

int
Summary::IsStgFile(void)
{
		m_hr = StgIsStorageFile(m_File);
		if( FAILED(m_hr) )
		{
			return(0);
		} 
		return(1);
	
}

int
Summary::IsWin2000OrNT(void)
{
	OSVERSIONINFO osvi;
	ZeroMemory(&osvi, sizeof(OSVERSIONINFO));
   	osvi.dwOSVersionInfoSize = sizeof(OSVERSIONINFO);
   	BOOL bOsVersionInfoEx;
   	if( !(bOsVersionInfoEx = GetVersionEx(&osvi)) )
	{
        	return(0);
   	}
	if(osvi.dwPlatformId == VER_PLATFORM_WIN32_NT)
		return(1);
	return(0);
}


int
Summary::IsNTFS(void)
{
	
	return(1);
}

void
Summary::PropertyPIDToCaption(PROPSPEC propspec, char *title)
{
	//printf("propspec.propid: %i\n", propspec.propid);
	char t[50] = { '\0' };
	sprintf(t, " propspec.propid %i ", propspec.propid);
	switch (propspec.propid)
	{
		case PID_CODEPAGE:
     			strcpy(title, "Codepage");
     			break;
		case PID_TITLE:
     			strcpy(title, "Title");
     			break;
   		case PID_SUBJECT:
     			strcpy(title, "Subject");
     			break;
   		case PID_AUTHOR:
     			strcpy(title, "Author");
     			break;
   		case PID_KEYWORDS:
     			strcpy(title, "Keywords");
     			break;
   		case PID_COMMENTS:
     			strcpy(title, "Comments");
     			break;
   		case PID_TEMPLATE:
     			strcpy(title, "Template");
     			break;
   		case PID_LASTAUTHOR:
     			strcpy(title, "Last Saved By");
     			break;
   		case PID_REVNUMBER:
     			strcpy(title, "Revision Number");
     			break;
   		case PID_EDITTIME:
     			strcpy(title, "Total Editing Time");
     			break;
   		case PID_LASTPRINTED:
     			strcpy(title, "Last Printed");
     			break;
   		case PID_CREATE_DTM:
     			strcpy(title, "Create Time/Date");
     			break;
   		case PID_LASTSAVE_DTM:
     			strcpy(title, "Last Saved Time/Date");
     			break;
   		case PID_PAGECOUNT:
     			strcpy(title, "Number of Pages");
     			break;
   		case PID_WORDCOUNT:
     			strcpy(title, "Number of Words");
     			break;
   		case PID_CHARCOUNT:
     			strcpy(title, "Number of Characters");
     			break;
   		case PID_THUMBNAIL:
     			strcpy(title, "Thumbnail");
     			break;
   		case PID_APPNAME:
     			strcpy(title, "Creating Application");
     			break;
   		case PID_SECURITY:
	     		strcpy(title, "Security");
     			break;
   		default:
     			strcpy(title, "Type not defined");
     }
	//printf("The title: %s + %s\n", title, t);
}

void
Summary::HrToString(HRESULT hr, char *string) {
			if(hr == S_OK)
				strcpy(string,"S_OK");
			if(hr == E_ACCESSDENIED)
				strcpy(string,"E_ACCESSDENIED");
			if(hr == E_FAIL)
				strcpy(string,"E_FAIL");
			if(hr == E_HANDLE)
				strcpy(string,"E_HANDLE");
			if(hr == E_INVALIDARG)
				strcpy(string,"E_INVALIDARG");
			if(hr == E_NOTIMPL)
				strcpy(string,"E_NOTIMPL");
			if(hr == E_OUTOFMEMORY)
				strcpy(string,"E_OUTOFMEMORY");
			if(hr == E_PENDING)
				strcpy(string,"E_PENDING");
			if(hr == E_POINTER)
				strcpy(string,"E_POINTER");
			if(hr == E_UNEXPECTED)
				strcpy(string,"E_UNEXPECTED");
			if(hr == S_FALSE)
				strcpy(string,"S_FALSE");
			if(hr == STG_E_INVALIDPOINTER)
				strcpy(string,"STG_E_INVALIDPOINTER");
			if(hr == STG_E_INVALIDPARAMETER)
				strcpy(string,"STG_E_INVALIDPARAMETER");
			if(hr == E_NOINTERFACE )
				strcpy(string,"E_NOINTERFACE");
			if(hr == STG_E_INVALIDFLAG )
				strcpy(string,"STG_E_INVALIDFLAG");
			if(hr == STG_E_INVALIDNAME )
				strcpy(string,"STG_E_INVALIDNAME");
			if(hr == STG_E_INVALIDFUNCTION )
				strcpy(string,"STG_E_INVALIDFUNCTION");
			if(hr == STG_E_LOCKVIOLATION )
				strcpy(string,"STG_E_LOCKVIOLATION");
			if(hr == STG_E_SHAREVIOLATION )
				strcpy(string,"STG_E_SHAREVIOLATION");
			if(hr == STG_E_UNIMPLEMENTEDFUNCTION )
				strcpy(string,"STG_E_UNIMPLEMENTEDFUNCTION");
			if(hr == STG_E_INCOMPLETE )
				strcpy(string,"STG_E_INCOMPLETE");
			if(hr == STG_E_ACCESSDENIED) 
				strcpy(string,"STG_E_ACCESSDENIED");
			if(hr == STG_E_FILENOTFOUND)
				strcpy(string,"STG_E_FILENOTFOUND");
}

SV* Summary::ReadOOo(SV* ptResult)
{
	printf("The document is from OpenOffice\n");
	return(ptResult);
}

SV* Summary::Read(void) {
		
	SV* ptResult = NEWSV(0,0);
	IPropertyStorage *pPropStg = NULL;
	char tmp1[1024];
	char tmp[1000] = { '\0' };
	char err[2];
	strcpy(err, "0");
	PROPVARIANT propvar;
	IEnumSTATPROPSTG *penum;
	STATPROPSTG PropStat;
        PROPSPEC propspec;
        SYSTEMTIME SystemTime;
	//if (m_IsOOo == 1)
	//{
	//	return this->ReadOOo(ptResult);
	//}
	m_hr = StgOpenStorageEx( m_File,
                          STGM_READ|STGM_SHARE_DENY_WRITE,
                          STGFMT_ANY,
                          0,
                          NULL,
                          NULL,
                           IID_IPropertySetStorage,
                           reinterpret_cast<void**>(&m_ipStg) );
	if( FAILED(m_hr) ) 
	{
		
		SetErr("could not open storage for inputfile: ");
		return(newRV_noinc(newSVpvn(err, strlen(err))));

	} 
	//m_hr = m_ipStg->Open(FMTID_UserDefinedProperties, STGM_READ|STGM_SHARE_EXCLUSIVE, &pPropStg);
	m_hr = m_ipStg->Open(FMTID_SummaryInformation, STGM_READ|STGM_SHARE_EXCLUSIVE, &pPropStg);
	if(FAILED(m_hr) )
	{
		SetErr("m_ipStg->Open failed: ");
		return(newRV_noinc(newSVpvn(err, strlen(err))));

	} 

	m_hr = pPropStg->Enum(&penum);
	if(FAILED(m_hr) )
	{
		SetErr("PropStg->Enum failed: ");
		return(newRV_noinc(newSVpvn(err, strlen(err))));
		
	} 
	m_hr = penum->Next(1, &PropStat, NULL);
	while(m_hr == S_OK)
	{
		propspec.ulKind = PRSPEC_PROPID;
		propspec.propid = PropStat.propid;
		//printf("Type: %d\n", PropStat.propid);
		PropVariantInit( &propvar );
		m_hr = pPropStg->ReadMultiple( 1, &propspec, &propvar );
		if( FAILED(m_hr) )
		{
			SetErr("pPropStg->ReadMultiple failed: ");
			return(newRV_noinc(newSVpvn(err, strlen(err))));
		}
		PropertyPIDToCaption(propspec, tmp1);
		
		//if(VT_I4  == propvar.vt) propspec.propid = 19
			//printf("propvar.vt %i!!!!!!!!!!!!!\n", propvar.vt);
		
		if(propvar.vt == VT_LPSTR)
		{
			if(hv_store(m_hv, tmp1, (U32) strlen(tmp1), newSVpv(propvar.pszVal, 0),0) == NULL)
				croak("Can not store in Hash!\n");
		} else if( propvar.vt == VT_FILETIME )
		{
			FileTimeToSystemTime(&propvar.filetime, &SystemTime);
			wsprintf(tmp, TEXT("%02d/%02d/%d  %02d:%02d:%02d"),
       					SystemTime.wMonth, SystemTime.wDay, SystemTime.wYear, SystemTime.wHour, SystemTime.wMinute, SystemTime.wSecond );
			if(hv_store(m_hv, tmp1, (U32) strlen(tmp1), newSVpv(tmp, 0),0) == NULL)
				croak("Can not store in Hash!\n");
			
		} else if( propvar.vt == VT_I4)
		{
			wsprintf(tmp, TEXT("%d"), propvar.lVal);
			if(hv_store(m_hv, tmp1, (U32) strlen(tmp1), newSVpv(tmp, 0),0) == NULL)
				croak("Can not store in Hash!\n");
		} else
		{
			printf("%s: ", tmp1);
			printf("Value: %d VT: %d\n",propvar.lVal,  propvar.vt);

		}
	
		if(PropStat.lpwstrName) {
			 CoTaskMemFree(PropStat.lpwstrName);
		}
		m_hr = penum->Next(1, &PropStat, NULL);
	} // end while m_hr == S_OK
	printf("Ende\n\n");
	ptResult = newRV_noinc((SV *) m_hv);
	return (ptResult);
}

int
Summary::Write(void)
{
	//m_hr = StgCreateStorageEx(
	//	m_File,
	//);
	return (1);
}

MODULE = Win32::File::Summary		PACKAGE = Win32::File::Summary		

Summary* 
Summary::new(File)
	char * File

MODULE = Win32::File::Summary		PACKAGE = SummaryPtr
	
SV*
Summary::Read()

int
Summary::Write()

SV*
Summary::GetError();
					
int
Summary::IsWin2000OrNT()

int
Summary::IsStgFile()

int
Summary::IsNTFS()

void
Summary::DESTROY()

int
Summary::IsOOoFile()