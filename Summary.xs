
//#ifdef __cplusplus
//extern "C" {
//#endif
#include "windows.h"
#include "Objidl.h"
#include "Objbase.h"
#include "defines.h"
#include "zip.h"
#include "unzip.h"
#include "OOo.h"

#undef THIS 

#include "EXTERN.h"
#include "perl.h"
#include "XSUB.h"
//#ifdef __cplusplus
//}
//#endif



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
			m_av = newAV();
			m_wFile=File;
			len = strlen(File);
			tmp[0] = File[len-3];
			tmp[1] = File[len-2];
			tmp[2] = File[len-1];
			tmp[3] = '\0';
			//printf("tmp : %s\n", tmp);
			m_IsOOo = 0;
			m_oemcp = 0;
			if(!strcmp(tmp,"sxw")) m_IsOOo = 1;
			if(!strcmp(tmp,"sxc")) m_IsOOo = 1;
			if(!strcmp(tmp,"ods")) m_IsOOo = 1;
			if(!strcmp(tmp,"odt")) m_IsOOo = 1;
			if(!strcmp(tmp,"odg")) m_IsOOo = 1;
			if(!strcmp(tmp,"odp")) m_IsOOo = 1;
		}
		~Summary();
		int IsWin2000OrNT(void);
		int IsStgFile(void);
		int IsNTFS(void);
		SV* Read(void);
		SV* Write(SV* newdata);
		SV* GetError(void);
		int IsOOoFile(void) { return(m_IsOOo); }
		void SetOEMCP(int oemcp) { m_oemcp=oemcp; }
		SV* _GetTitles(void);
	private:
		void HrToString(HRESULT hr, char *string);
		void PropertyPIDToCaption(PROPSPEC propspec, char *title);
		void SetErr(char *msg);
		void ReadOOo(void);
		void ReadXML(void);
	private:
		wchar_t m_File[2048];
		char * m_wFile;
		char m_perror[1024];
		HRESULT m_hr;
		IPropertySetStorage *m_ipStg;
		HV* m_hv;
		AV* m_av;
		int m_IsOOo;	// 1 if the file is an OpenOffice document
		bool m_oemcp;
		//SV* m_ptResult;
		
		

};

SV*
Summary::_GetTitles(void)
{
	return newRV((SV *)m_av);
}

void
Summary::ReadXML(void)
{
	if(hv_store(m_hv, "Not jet implemented", (U32) strlen("Not jet implemented"), newSVpv("0", 0),0) == NULL)
		croak("Can not store in Hash!\n");
	
}
	
void
Summary::ReadOOo(void)
{
	unzFile uzFile;
	unz_global_info pglobal_info;
	uInt size_buff = 10000;
	char *buff = (char *)malloc(size_buff+1);
	uzFile = unzOpen(m_wFile);
	if(uzFile != NULL)
	{
		if(unzGetGlobalInfo(uzFile, &pglobal_info) != UNZ_OK)
			croak("Error in unzGetGlobalInfo!\n");

		if(unzLocateFile(uzFile,"meta.xml", 2) != UNZ_OK)
			croak("can not locate meta.xml!\n");

		if(unzOpenCurrentFile(uzFile) != UNZ_OK)
			croak("can not open meta.xml!\n");

		if(unzReadCurrentFile(uzFile, buff, size_buff) < 0)
			croak("can not read meta.xml!\n");
		
		OOo *OOo_data = new OOo;	// initializeing 
		
		OOo_data->SetBuffer(buff,m_oemcp);	// setting the buffer
		
		if(OOo_data->ParseBuffer(m_hv,m_av) == false)
		{
			printf("Error parsing XML - buffer!\n");
		}
		free(buff);	// Freeing temp. buffer
		unzCloseCurrentFile(uzFile);
		unzClose(uzFile);
	} else
	{
		printf("could not open OpenOffice/StarOffice Document!\n");
	}
	
}

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
}

Summary::~Summary()
{ 
	if( m_ipStg ) m_ipStg->Release();
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
	//printf("propspec.ulKind : %i propspec.propid %0x\n", propspec.ulKind, propspec.propid );
	if(propspec.ulKind == 1) {
		switch (propspec.propid)
		{
			case PID_CODEPAGE:
     				strcpy(title, "Codepage");
     				break;
			case PIDSI_TITLE:
     				strcpy(title, "Title");
     				break;
   			case PIDSI_SUBJECT:
     				strcpy(title, "Subject");
     				break;
   			case PIDSI_AUTHOR:
     				strcpy(title, "Author");
     				break;
   			case PIDSI_KEYWORDS:
     				strcpy(title, "Keywords");
     				break;
   			case PIDSI_COMMENTS:
     				strcpy(title, "Comments");
     				break;
   			case PIDSI_TEMPLATE:
     				strcpy(title, "Template");
     				break;
   			case PIDSI_LASTAUTHOR:
     				strcpy(title, "Last Saved By");
     				break;
   			case PIDSI_REVNUMBER:
     				strcpy(title, "Revision Number");
     				break;
   			case PIDSI_EDITTIME:
     				strcpy(title, "Total Editing Time");
     				break;
   			case PIDSI_LASTPRINTED:
     				strcpy(title, "Last Printed");
     				break;
   			case PIDSI_CREATE_DTM:
     				strcpy(title, "Create Time/Date");
     				break;
   			case PIDSI_LASTSAVE_DTM:
     				strcpy(title, "Last Saved Time/Date");
     				break;
   			case PIDSI_PAGECOUNT:
     				strcpy(title, "Number of Pages");
     				break;
   			case PIDSI_WORDCOUNT:
     				strcpy(title, "Number of Words");
     				break;
   			case PIDSI_CHARCOUNT:
     				strcpy(title, "Number of Characters");
     				break;
   			case PIDSI_THUMBNAIL:
     				strcpy(title, "Thumbnail");
     				break;
   			case PIDSI_APPNAME:
     				strcpy(title, "Creating Application");
     				break;
   			case PIDSI_DOC_SECURITY:
		     		strcpy(title, "Security");
     				break;
   			default:
     				strcpy(title, "Type not defined");
     		}
     	} else
     		strcpy(title, (char*)propspec.lpwstr );
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
				strcpy(string,"The requested storage does not exist (STG_E_FILENOTFOUND).");
}


SV* Summary::Read(void) {
		
	SV* m_ptResult = NEWSV(0,0);
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
	if (m_IsOOo == 1)
	{
		
		this->ReadOOo();
	}
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
		if(m_IsOOo == 0) {
			SetErr("m_ipStg->Open failed: ");
			return(newRV_noinc(newSVpvn(err, strlen(err))));
		} else
		{
			m_ptResult = newRV_noinc((SV *) m_hv);
			return (m_ptResult);
		}
	} 

	m_hr = pPropStg->Enum(&penum);
	if(FAILED(m_hr) )
	{
		if(m_IsOOo == 0) {
			SetErr("PropStg->Enum failed: ");
			return(newRV_noinc(newSVpvn(err, strlen(err))));
		} else
		{
			m_ptResult = newRV_noinc((SV *) m_hv);
			return (m_ptResult);
		}
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
		av_push(m_av, newSVpvn(tmp1, strlen(tmp1)));
		
		//if(VT_I4  == propvar.vt) propspec.propid = 19
			//printf("propvar.vt %d !!!!!!!!!!!!! %s\n", propvar.vt, tmp1);
		
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
	m_ptResult = newRV_noinc((SV *) m_hv);
	return (m_ptResult);
}

SV* Summary::Write(SV* newdata)
{
	HV* newinfo = (HV*) SvRV(newdata);
	m_hr = StgOpenStorageEx( m_File,
                          STGM_WRITE|STGM_SHARE_DENY_WRITE ,
                          STGFMT_ANY,
                          0,
                          NULL,
                          NULL,
                           IID_IPropertySetStorage,
                           reinterpret_cast<void**>(&m_ipStg) );
	if( FAILED(m_hr) ) 
	{
		SetErr("could not open storage for inputfile: ");
		return(newRV_noinc(newSVpvn("0", strlen("0"))));

	} 
	return (newRV_noinc(newSVpvn("1", strlen("1"))));
}

MODULE = Win32::File::Summary		PACKAGE = Win32::File::Summary		

Summary* 
Summary::new(File)
	char * File

SV*
Summary::_GetTitles()

	
SV*
Summary::Read()

void 
Summary::SetOEMCP(oemcp)
		int oemcp

SV*
Summary::Write(newdata)
			SV* newdata

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

