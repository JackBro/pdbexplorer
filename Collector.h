
#pragma once


typedef struct tagPDBSYMBOL {
   CString sKey;
   DWORD dwSymId;
   DWORD dwAddrSect;
   DWORD dwAddrOffset;
   DWORD dwSymTag;
   DWORD dwUniqueId;
} PDBSYMBOL;


class CPdbCollector : public CThreadImpl<CPdbCollector>
{
public:
   HWND m_hWnd;
   CString m_sFilename;
   CString m_sLastError;
   CComAutoCriticalSection m_lock;
   IDiaDataSourcePtr m_spDataSource;
   CAtlArray<PDBSYMBOL> m_aSymbols;
   volatile bool m_bDone;

   CPdbCollector() : m_hWnd(NULL), m_bDone(false)
   {
   }

   void Init(HWND hWnd, LPCTSTR pstrFilename)
   {
      m_aSymbols.RemoveAll();
      m_sFilename = pstrFilename;
      m_sLastError.Empty();
      m_hWnd = hWnd;
      m_bDone = false;
   }

   DWORD Run()
   {
      ::CoInitialize(NULL);
      ::PostMessage(m_hWnd, WM_USER_LOAD_START, 0, 0L);
      m_bDone = false;
      _ATLTRY
      {
         m_spDataSource.CreateInstance(CLSID_DiaSource);
         m_spDataSource->loadDataFromPdb(CT2W(m_sFilename));
         IDiaSessionPtr spSession;
         m_spDataSource->openSession(&spSession);
         if( IsAborted() ) _Abort();
         _ProcessTables(spSession);
      }
      _ATLCATCHALL()
      {
         ::PostMessage(m_hWnd, WM_USER_LOAD_END, 0, 0L);
      }
      m_spDataSource.Release();
      m_bDone = true;
      ::CoUninitialize();
      return 0;
   }

   void _ProcessTables(IDiaSessionPtr& spSession) throw(...)
   {
      IDiaEnumTablesPtr spEnumTables;
      spSession->getEnumTables(&spEnumTables);
      IDiaTablePtr spTable;
      ULONG nCount = 0;
      while( SUCCEEDED( spEnumTables->Next(1, &spTable, &nCount) ) && nCount == 1 ) {
         CComBSTR bstrTableName;
         spTable->get_name(&bstrTableName);
         IDiaEnumSymbolsPtr spEnumSymbols;
         IDiaEnumSourceFilesPtr spEnumSourceFiles;
         spEnumSymbols = spTable;
         spEnumSourceFiles = spTable;
         if( spEnumSourceFiles != NULL ) {
            IDiaSourceFilePtr spSourceFile;
            while( SUCCEEDED( spEnumSourceFiles->Next(1, &spSourceFile, &nCount) ) && nCount == 1 ) {
               _ProcessSourceFile(bstrTableName, spSourceFile);
               nCount = 0;
               spSourceFile.Release();
            }
         }
         if( spEnumSymbols != NULL ) {
            IDiaSymbolPtr spSymbol;
            while( SUCCEEDED( spEnumSymbols->Next(1, &spSymbol, &nCount) ) && nCount == 1 ) {
               _ProcessSymbol(bstrTableName, spSymbol);
               nCount = 0;
               spSymbol.Release();
            }
         }
         spTable.Release();
      }
   }

   void _ProcessSourceFile(LPCWSTR pwstrTableName, IDiaSourceFilePtr& spSourceFile)
   {
      DWORD dwUniqueId = 0;
      spSourceFile->get_uniqueId(&dwUniqueId);
      CString sKey;
      sKey.Format(_T("Source Files|%ls"), pwstrTableName);
      CComCritSecLock<CComCriticalSection> lock(m_lock);
      PDBSYMBOL Info = { sKey, 0, 0, 0, SymTagCompiland, dwUniqueId };
      m_aSymbols.Add(Info);
      if( IsAborted() ) throw 1;
   }

   void _ProcessSymbol(LPCWSTR pwstrTableName, IDiaSymbolPtr& spSymbol)
   {
      DWORD dwSymTag = SymTagNull;
      spSymbol->get_symTag(&dwSymTag);
      LPCTSTR pstrScope = _T("Misc");
      switch( dwSymTag ) {
      case SymTagExe:          pstrScope = _T("Executables"); break;
      case SymTagCompiland:    pstrScope = _T("Object Files"); break;         
      case _SymTagEnum:        pstrScope = _T("Enumerations"); break;
      case SymTagFunctionType: pstrScope = _T("Function Types"); break;
      case SymTagTypedef:      pstrScope = _T("Typedefs"); break;
      case SymTagUDT:          pstrScope = _T("Types"); break;
      case SymTagFunction:     pstrScope = _T("Private Classes"); break;         
      case SymTagPublicSymbol: pstrScope = _T("Public Classes"); break;
         // Symbols we don't care about
      case SymTagFuncDebugStart:
      case SymTagFuncDebugEnd:
      case SymTagVTable:
      case SymTagThunk:
      case SymTagLabel:
      case SymTagData:
         return;
      }
      CComBSTR bstrName;
      spSymbol->get_undecoratedNameEx(0x1000, &bstrName);
      if( bstrName.m_str == NULL ) spSymbol->get_name(&bstrName);
      if( bstrName.m_str == NULL ) spSymbol->get_undecoratedName(&bstrName);
      if( bstrName.m_str == NULL ) spSymbol->get_libraryName(&bstrName);
      ULONG lSymbolID = 0;
      spSymbol->get_symIndexId(&lSymbolID);
      CString sName = bstrName;
      if( dwSymTag == SymTagPublicSymbol && sName.Find(_T("::")) < 0 ) pstrScope = _T("Public Globals");
      if( dwSymTag == SymTagFunction && sName.Find(_T("::")) < 0 ) pstrScope = _T("Private Globals");
      if( sName.Find('@') >= 0 ) sName = sName.Left(sName.Find('@') - 1);
      if( sName.Find('\\') >= 0 ) sName = sName.Right(sName.GetLength() - sName.ReverseFind('\\') - 1);
      CString sKeyName = sName;
      _StripCppName(sKeyName);
      CString sKey;
      sKey.Format(_T("%ls|%s|%s"), pwstrTableName, pstrScope, sKeyName);
      DWORD dwSymId = 0;
      DWORD dwAddrSect = 0;
      DWORD dwAddrOffset = 0;
      spSymbol->get_symIndexId(&dwSymId);
      spSymbol->get_addressSection(&dwAddrSect);
      spSymbol->get_addressOffset(&dwAddrOffset);
      CComCritSecLock<CComCriticalSection> lock(m_lock);
      PDBSYMBOL Info = { sKey, dwSymId, dwAddrSect, dwAddrOffset, dwSymTag, 0 };
      m_aSymbols.Add(Info);
      if( IsAborted() ) throw 1;
   }

   void _Fail()
   {
      CComBSTR bstrError;
      m_spDataSource->get_lastError(&bstrError);
      m_sLastError = bstrError;
      m_spDataSource.Release();
   }

   void _StripCppName(CString& sKeyName)
   {
      int cch = sKeyName.GetLength();
      int iParen = 0, iBracket = 0;
      for( int i = 0; i < cch; i++ ) {
         switch( sKeyName[i] ) {
         case '<': iBracket++; break;
         case '>': iBracket--; break;
         case '(': iParen++; break;
         case ')': iParen--; break;
         case ':': if( iBracket > 0 || iParen > 0 ) sKeyName.SetAt(i, 'Ø'); break;
         }
      }
      sKeyName.Replace('|', '!');
      sKeyName.Replace(_T("::"), _T("|"));
      sKeyName.Replace('Ø', ':');
      sKeyName.Replace(_T("__sz_"), _T(""));
      sKeyName.Replace(_T("__imp_load_"), _T(""));
      sKeyName.Replace(_T("__imp_"), _T(""));
      sKeyName.Replace(_T("`"), _T(""));
      sKeyName.Replace(_T("'"), _T(""));
   }
};

