
#pragma once


class CDiaInfo
{
public:
   CString GetSymbolInfo(IDiaDataSourcePtr& spDataSource, const PDBSYMBOL& Symbol)
   {
      CString sRTF;
      _AddHeader(sRTF);
      _ATLTRY
      {
         HRESULT Hr;

         IDiaSessionPtr spSession;
         spDataSource->openSession(&spSession);
         IDiaSymbolPtr spSymbol;
         IDiaSymbolPtr spGlobal;
         IDiaSourceFilePtr spSourceFile;
         spSession->get_globalScope(&spGlobal);
         spSession->symbolById(Symbol.dwSymId, &spSymbol);

         if( spSymbol == NULL ) spSession->findSymbolByAddr(Symbol.dwAddrSect, Symbol.dwAddrOffset, (SymTagEnum) Symbol.dwSymTag, &spSymbol);

         if( spSymbol == NULL ) spSession->findFileById(Symbol.dwUniqueId, &spSourceFile);

         if( spSymbol == NULL ) {
            IDiaEnumSymbolsPtr spChildren;
            spGlobal->findChildren(SymTagNull, CT2W(Symbol.sKey.Mid(Symbol.sKey.ReverseFind('|') + 1)), 0, &spChildren);
            ULONG celt = 0;
            if( spChildren != NULL ) spChildren->Next(1, &spSymbol, &celt);
         }

         if( spSymbol == NULL ) spSymbol = spSourceFile;

         CComBSTR bstrTitle;
         if( bstrTitle.m_str == NULL ) spSymbol->get_undecoratedNameEx(0x1000, &bstrTitle);
         if( bstrTitle.m_str == NULL ) spSymbol->get_undecoratedName(&bstrTitle);
         if( bstrTitle.m_str == NULL ) spSymbol->get_name(&bstrTitle);
         if( bstrTitle.m_str == NULL ) spSymbol->get_libraryName(&bstrTitle);

         long bIsFunction = FALSE;
         long bIsConstructor = FALSE;
         spSymbol->get_function(&bIsFunction);
         spSymbol->get_constructor(&bIsConstructor);

         DWORD dwSymTag = 0;
         spSymbol->get_symTag(&dwSymTag);
         LPCTSTR pstrSymTags[] = { _T("NULL"), _T("Executable"), _T("Compiland"), _T("Compiland Details"), _T("Compiland Env"), _T("Function"), _T("Block"), _T("Data"), _T("Annotation"), _T("Label"), _T("Public Symbol"), _T("User Defined Type"), _T("Enumeration"), _T("Function Type"), _T("Pointer Type"), _T("Array Type"), _T("Base Type"), _T("Typedef"), _T("Class"), _T("Friend"), _T("Argument Type"), _T("Debug Start"), _T("Debug End"), _T("Using Namespace"), _T("VTable Shape"), _T("VTable"), _T("Custom"), _T("Thunk"), _T("Custom Type"), _T("Managed Type"), _T("Dimension") };
         CString sType = _T("Misc");
         if( dwSymTag < _countof(pstrSymTags) ) sType = pstrSymTags[dwSymTag];
         if( bIsFunction ) sType = _T("Function");
         if( bIsConstructor ) sType = _T("Constructor / Destructor");
         _AddTitle(sRTF, bstrTitle.m_str, sType);

         CComBSTR bstrUndecorated;
         spSymbol->get_undecoratedNameEx(0, &bstrUndecorated);
         if( bstrUndecorated.Length() > 0 ) {
            sRTF += _T("\\fs22\\b Signature:\\par\\fs24\\b0 ");
            sRTF += bstrUndecorated;
            sRTF += _T(" \\par\\par ");
         }

         DWORD dwDataKind = 0;
         Hr = spSymbol->get_dataKind(&dwDataKind);
         LPCTSTR pstrDataKinds[] = { _T(""), _T("Local"), _T("Static Local"), _T("Parameter"), _T("Object Pointer"), _T("File Static"), _T("Global"), _T("Member"), _T("Static Member"), _T("Constant") };
         CString sDataKind;
         if( dwDataKind < _countof(pstrDataKinds) ) sDataKind = pstrDataKinds[dwDataKind];
         if( Hr == S_OK ) _AddTableEntry(sRTF, _T("Data Kind"), sDataKind);

         GUID guid = { 0 };
         spSymbol->get_guid(&guid);
         if( guid != GUID_NULL ) {
            CComBSTR bstrGUID(guid);
            _AddTableEntry(sRTF, _T("GUID"), bstrGUID.m_str);
         }

         DWORD dwLanguage = 0;
         Hr = spSymbol->get_language(&dwLanguage);
         LPCTSTR pstrLanguages[] = { _T("C"), _T("C++"), _T("Fortran") _T("MASM"), _T("Pascal"), _T("Basic"), _T("Cobol"), _T("Link"), _T("Cvtres"), _T("Cvtpgd"), _T("C#"), _T("VB"), _T("ILASM"), _T("Java"), _T("JavaScript"), _T("MSIL") };
         CString sLanguage;
         if( dwLanguage < _countof(pstrLanguages) ) sLanguage = pstrLanguages[dwLanguage];
         if( Hr == S_OK ) _AddTableEntry(sRTF, _T("Language"), sLanguage);

         CComBSTR bstrSourceFile;
         spSymbol->get_sourceFileName(&bstrSourceFile);
         if( bstrSourceFile.m_str == NULL ) spSymbol->get_symbolsFileName(&bstrSourceFile);
         _AddTableEntry(sRTF, _T("Source File"), bstrSourceFile.m_str);

         CComBSTR bstrCompiler;
         spSymbol->get_compilerName(&bstrCompiler);
         _AddTableEntry(sRTF, _T("Compiler"), bstrCompiler.m_str);

         CComVariant vValue;
         Hr = spSymbol->get_value(&vValue);
         if( Hr == S_OK ) {
            if( SUCCEEDED( vValue.ChangeType(VT_BSTR) ) ) _AddTableEntry(sRTF, _T("Value"), vValue.bstrVal);
         }

         ULONG nTypes = 0;
         IDiaSymbol* ppTypeSymbols[10] = { 0 };
         spSymbol->get_types(_countof(ppTypeSymbols), &nTypes, ppTypeSymbols);
         CString sTypes;
         for( ULONG i = 0; i < nTypes && ppTypeSymbols[i] != NULL; i++ ) {
            CComBSTR bstrName;
            ppTypeSymbols[i]->get_name(&bstrName);
            if( !sTypes.IsEmpty() ) sTypes += _T(", "); sTypes += bstrName;
            ppTypeSymbols[i]->Release();
         }
         _AddTableEntry(sRTF, _T("Type"), sTypes);

         IDiaEnumSymbolsPtr spChildren;
         spSymbol->findChildren(SymTagNull, NULL, 0, &spChildren);
         if( spChildren != NULL ) {
            ULONG celt = 0;
            CString sChildren;
            IDiaSymbolPtr spChildSymbol;
            while( spChildren->Next(1, &spChildSymbol, &celt), celt > 0 ) {
               DWORD dwSymTag = 0;
               spChildSymbol->get_symTag(&dwSymTag);
               if( dwSymTag != SymTagBaseClass && dwSymTag != SymTagBaseType ) continue;
               CComBSTR bstrName;
               spChildSymbol->get_name(&bstrName);
               if( !sChildren.IsEmpty() ) sChildren += _T(", "); sChildren += bstrName;
               spChildSymbol.Release();
            }
            _AddTableEntry(sRTF, _T("Lex"), sChildren);
         }

         IDiaSymbolPtr spBaseTypeSymbol;
         spSymbol->get_type(&spBaseTypeSymbol);
         if( spBaseTypeSymbol != NULL ) {
            CComBSTR bstrName;
            spBaseTypeSymbol->get_name(&bstrName);
            _AddTableEntry(sRTF, _T("Base Type"), bstrName.m_str);
         }

         IDiaSymbolPtr spObjectTypeSymbol;
         spSymbol->get_objectPointerType(&spObjectTypeSymbol);
         if( spObjectTypeSymbol != NULL ) {
            CComBSTR bstrName;
            spObjectTypeSymbol->get_name(&bstrName);
            _AddTableEntry(sRTF, _T("Object Type"), bstrName.m_str);
         }

         IDiaSymbolPtr spContainerSymbol;
         spSymbol->get_container(&spContainerSymbol);
         if( spContainerSymbol != NULL ) {
            CComBSTR bstrName;
            spContainerSymbol->get_name(&bstrName);
            _AddTableEntry(sRTF, _T("Parent Type"), bstrName.m_str);
         }

         IDiaSymbolPtr spBaseClassSymbol;
         spSymbol->get_virtualBaseTableType(&spBaseClassSymbol);
         if( spBaseClassSymbol != NULL ) {
            CComBSTR bstrName;
            spBaseClassSymbol->get_name(&bstrName);
            _AddTableEntry(sRTF, _T("Base Class"), bstrName.m_str);
         }

         IDiaSymbolPtr spBaseShapeSymbol;
         spSymbol->get_virtualTableShape(&spBaseShapeSymbol);
         if( spBaseShapeSymbol != NULL ) {
            CComBSTR bstrName;
            spBaseShapeSymbol->get_name(&bstrName);
            _AddTableEntry(sRTF, _T("Base Class"), bstrName.m_str);
         }

         IDiaSymbolPtr spClassParentSymbol;
         spSymbol->get_classParent(&spClassParentSymbol);
         if( spClassParentSymbol != NULL ) {
            CComBSTR bstrName;
            spClassParentSymbol->get_name(&bstrName);
            _AddTableEntry(sRTF, _T("Class Parent"), bstrName.m_str);
         }

         IDiaSymbolPtr spLexParentSymbol;
         spSymbol->get_lexicalParent(&spLexParentSymbol);
         if( spLexParentSymbol != NULL ) {
            CComBSTR bstrName;
            spLexParentSymbol->get_name(&bstrName);
            _AddTableEntry(sRTF, _T("Owner"), bstrName.m_str);
         }
      }
      _ATLCATCHALL()
      {
      }
      _AddFooter(sRTF);
      return sRTF;
   }

   CString GetEmptyInfo(CTreeViewCtrl& ctrlTree) const
   {
      CString sRTF, sTitle;
      ctrlTree.GetItemText(ctrlTree.GetSelectedItem(), sTitle);
      _AddHeader(sRTF);
      _AddTitle(sRTF, sTitle, _T("Group"));
      _AddFooter(sRTF);
      return sRTF;
   }

   // Implementation

   void _AddHeader(CString& sRTF) const
   {
      sRTF += _T("{\\rtf1\\ansi\\ansicpg1252\\deff0\\deflang1030{\\fonttbl{\\f0\\fnil\\fcharset0 Arial;}} {\\viewkind4\\uc1\\pard\\sa200\\sl236\\slmult1\\li100 ");
   }

   void _AddTitle(CString& sRTF, CString sTitle, CString sGroup) const
   {
      sRTF += _T(" \\par\\f0\\fs52 ");
      sRTF += sTitle;
      sRTF += _T(" \\par\\sl1000\\fs30 ");
      sRTF += sGroup;
      sRTF += _T(" \\par\\par ");
   }

   void _AddFooter(CString& sRTF) const
   {
      sRTF += _T("\\par\\par}");
   }

   void _AddTableEntry(CString& sRTF, CString sLabel, CString sValue)
   {
      if( sValue.IsEmpty() ) return;
      sValue.Replace(_T("\\"), _T("\\\\"));
      sValue.Replace(_T("{"), _T("\\{ "));
      CString sTemp;
      sTemp.Format(_T("\\fs20\\b %s: \\b0 %s \\par "), sLabel, sValue);
      sRTF += sTemp;
   }
};

