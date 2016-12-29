// ADOSample.cpp : 定義主控台應用程式的進入點。
//

#include "stdafx.h"
#include <ole2.h>
#include <stdio.h>

#import "msado15.dll" no_namespace rename("EOF", "EndOfFile")

inline void TESTHR(HRESULT x) { if FAILED(x) _com_issue_error(x); };
void ExecuteX();
void PrintProviderError(_ConnectionPtr pConnection);
void PrintComError(_com_error &e);
void PrintOutput(_RecordsetPtr pRstTemp, int type = 1);
void ExecuteCommand(_CommandPtr pCmdTemp, _RecordsetPtr pRstTemp);

int _tmain(int argc, _TCHAR* argv[])
{
	if ( FAILED(::CoInitialize(NULL)) ){
		return -1;
	}

	ExecuteX();
   ::CoUninitialize();
	return 0;
}

void ExecuteX() 
{

	// Define string variables.
	_bstr_t strSQLQuery("Select * from logmsg where SEQ = 1;");
	_bstr_t strCnn("Provider='sqloledb';Data Source=localhost; Initial Catalog='SDVR';UID=sa;PWD=p@ssw0rd;");
	//_bstr_t strCnn("Server=localhost;Database=DVR;User Id=sa;Password=p@ssw0rd;");
	

	// Define ADO object pointers.  Initialize pointers on define.
   // These are in the ADODB::  namespace.
	_ConnectionPtr pConnection = NULL;
	_CommandPtr pCmdChange = NULL;
	_RecordsetPtr pRstTitles = NULL;

	try {
		// Open connection.
		TESTHR(pConnection.CreateInstance(__uuidof(Connection)));
		pConnection->Open (strCnn, "", "", adConnectUnspecified);

		// Create command object.
		TESTHR(pCmdChange.CreateInstance(__uuidof(Command)));
		pCmdChange->ActiveConnection = pConnection;
		pCmdChange->CommandText = strSQLQuery;

		// Open titles table, casting Connection pointer to an 
		// IDispatch type so converted to correct type of variant.
		TESTHR(pRstTitles.CreateInstance(__uuidof(Recordset)));

		//pRstTitles->Open("select a.cust_no, a.cust_sna, b.VER_NO, c.IS_TEST, d.STA_YN, e.EXT_IP, e.EXT_PORT from customer a join cust_params b on a.cust_no=b.cust_no and b.CONNECTED='Y' and b.SENDER_TYPE='S' join cust_alias c on a.cust_no=c.cust_no join sys_status d on a.cust_no=d.cust_no and d.STA_NO='S01' join cust_comm e on a.cust_no=e.cust_no", _variant_t((IDispatch *)pConnection), adOpenStatic, adLockReadOnly, adCmdText);
		pRstTitles->Open("select HOT_NO from gogomedia Where cust_no='a000' and spd_no='1';", _variant_t((IDispatch *)pConnection), adOpenStatic, adLockReadOnly, adCmdText);
		//PrintOutput(pRstTitles);

		if ((pRstTitles == NULL) || (pRstTitles->RecordCount == 0)){
			printf("Trace1\n");
		}

		pRstTitles->MoveFirst();
		
		if (pRstTitles->EndOfFile) {
			printf("Trace2\n");
		} else {
			_variant_t var = pRstTitles->GetCollect("HOT_NO");
			if (var.vt != VT_NULL) {
				int retval = var;
				printf("Trace4 retval = %d\n", retval);
			} else {
				printf("Trace3\n");
			}
			//retval = g_dbparam.m_pRstTitles->Fields->GetItem(field)->Value;
		}


		if (pRstTitles) {
			if (pRstTitles->State == adStateOpen) {
				pRstTitles->Close();
			}
		}
		printf("\t====================\n");

		TESTHR(pRstTitles.CreateInstance(__uuidof(Recordset)));
		pRstTitles->Open ("logmsg", _variant_t((IDispatch *) pConnection, true), adOpenStatic, adLockOptimistic, adCmdTable);

		// Print report of original data.
		printf("\n\nData in Titles table before executing the query: \n");

		// Call function to print loop recordset contents.
		PrintOutput(pRstTitles, 0);
		
		// Clear extraneous errors from the Errors collection.
		pConnection->Errors->Clear();

		// Call ExecuteCommand subroutine to execute pCmdChange command.
		ExecuteCommand(pCmdChange, pRstTitles);
		
		// Print report of new data.
		printf("\n\n\tData in Titles table after executing the query: \n");
		PrintOutput(pRstTitles, 0);
		
		// Use Connection object's Execute method to execute SQL statement to restore data.
		pConnection->Execute(strSQLQuery, NULL, adExecuteNoRecords);

		// Retrieve the current data by requerying the recordset.
		pRstTitles->Requery(adCmdUnknown);

		// Print report of restored data.
		printf("\n\n\tData after exec. query to restore original info: \n");
		PrintOutput(pRstTitles, 0);

		if (pRstTitles) {
			if (pRstTitles->State == adStateOpen) {
				pRstTitles->Close();
			}
		}
		pRstTitles->Open("Delete Event_log Where LOG_DAT='2011/08/10'", _variant_t((IDispatch *)pConnection), adOpenStatic, adLockReadOnly, adCmdText);

	}
	catch (_com_error &e) {
      PrintProviderError(pConnection);
      PrintComError(e);
	}
	// Clean up objects before exit.
	if (pRstTitles) {
		if (pRstTitles->State == adStateOpen) {
			pRstTitles->Close();
		}
	}
	if (pConnection) {
		if (pConnection->State == adStateOpen) {
			pConnection->Close();
		}
	}

}


void PrintProviderError(_ConnectionPtr pConnection)
{
	// Print Provider Errors from Connection object.
	// pErr is a record object in the Connection's Error collection.
	ErrorPtr pErr = NULL;

	if ( (pConnection->Errors->Count) > 0 ) {
		long nCount = pConnection->Errors->Count;
		// Collection ranges from 0 to nCount -1.
		for ( long i = 0 ; i < nCount ; i++ ) {
			pErr = pConnection->Errors->GetItem(i);
			printf("\t Error number: %x\t%s", pErr->Number, pErr->Description);
		}
	}
}

void PrintComError(_com_error &e)
{
	_bstr_t bstrSource(e.Source());
	_bstr_t bstrDescription(e.Description());

	// Print Com errors.
	printf("Error\n");
	printf("\tCode = %08lx\n", e.Error());
	printf("\tCode meaning = %s\n", e.ErrorMessage());
	printf("\tSource = %s\n", (LPCSTR) bstrSource);
	printf("\tDescription = %s\n", (LPCSTR) bstrDescription);
}

void PrintOutput(_RecordsetPtr pRstTemp, int type)
{
	// Ensure at top of recordset.
	pRstTemp->MoveFirst();

	// If EOF is true, then no data and skip print loop.
	if ( pRstTemp->EndOfFile ){
		printf("\tRecordset empty\n");
	} else {
		// Define strings for output conversions.  Initialize to first record's values.
		_bstr_t bstrTitle;
		_bstr_t bstrType;

		// Enumerate Recordset and print from each.
		while ( !(pRstTemp->EndOfFile) ) {
			// Convert variant string to convertable string type.
			if (type == 0) {
				bstrTitle = pRstTemp->Fields->GetItem("SEQ")->Value;
				bstrType  = pRstTemp->Fields->GetItem("MSG")->Value;
			} else {
				bstrTitle = pRstTemp->Fields->GetItem("CUST_NO")->Value;
				bstrType  = pRstTemp->Fields->GetItem("SPD_NO")->Value;
			}
			printf("\t%s, %s \n", (LPCSTR) bstrTitle, (LPCSTR) bstrType);

			pRstTemp->MoveNext();
		}
	}

}

void ExecuteCommand(_CommandPtr pCmdTemp, _RecordsetPtr pRstTemp)
{
	try {
		// CommandText property already set before function was called.
		pCmdTemp->Execute(NULL, NULL, adCmdText);

		// Retrieve the current data by requerying the recordset.
		pRstTemp->Requery(adCmdUnknown);
	}

	catch(_com_error &e) {
		// Notify user of any errors that result from executing the query.
		// Pass a connection pointer accessed from the Recordset.
		PrintProviderError(pRstTemp->GetActiveConnection());
		PrintComError(e);
	}
}