/****************************** Module Header ******************************\
Module Name:  RecipePreviewHandler.cpp
Project:      CppShellExtPreviewHandler
Copyright (c) Microsoft Corporation.



This source is subject to the Microsoft Public License.
See http://www.microsoft.com/opensource/licenses.mspx#Ms-PL.
All other rights reserved.

THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, 
EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED 
WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
\***************************************************************************/

#include "RecipePreviewHandler.h"
#include <Shlwapi.h>
#include <Wincrypt.h>   // For CryptStringToBinary.
#include <msxml6.h>
#include <WindowsX.h>
#include <assert.h>
#include <atlstr.h>
#include <iostream>

#include "resource.h"

#pragma comment(lib, "Shlwapi.lib")
#pragma comment(lib, "Crypt32.lib")
#pragma comment(lib, "msxml6.lib")

#include <sstream>

#undef max
#undef min
#include <limits>
#include <arrow/api.h>
#include <arrow/filesystem/filesystem.h>
#include <arrow/ipc/writer.h>
#include <arrow/util/iterator.h>
#include <arrow/ipc/reader.h>
#include <arrow/ipc/api.h>


extern HINSTANCE   g_hInst;
extern long g_cDllRef;

static const int CorrectWidth = 25, CorrectHeight = 25;

inline int RECTWIDTH(const RECT &rc)
{
    return (rc.right - rc.left);
}

inline int RECTHEIGHT(const RECT &rc )
{
    return (rc.bottom - rc.top);
}


RecipePreviewHandler::RecipePreviewHandler() : m_cRef(1), m_pPathFile(NULL), 
    m_hwndParent(NULL), m_hwndPreview(NULL), m_punkSite(NULL)
{
    InterlockedIncrement(&g_cDllRef);
}

RecipePreviewHandler::~RecipePreviewHandler()
{
    if (m_hwndPreview)
    {
        DestroyWindow(m_hwndPreview);
        m_hwndPreview = NULL;
    }
    if (m_punkSite)
    {
        m_punkSite->Release();
        m_punkSite = NULL;
    }
    if (m_pPathFile)
    {
        LocalFree(m_pPathFile);
        m_pPathFile = NULL;
    }

    for (auto str : stored)
        free(str);
    stored.clear();

    InterlockedDecrement(&g_cDllRef);
}


#pragma region IUnknown

// Query to the interface the component supported.
IFACEMETHODIMP RecipePreviewHandler::QueryInterface(REFIID riid, void **ppv)
{
    static const QITAB qit[] = 
    {
        QITABENT(RecipePreviewHandler, IPreviewHandler),
    //    QITABENT(RecipePreviewHandler, IInitializeWithStream), 
        QITABENT(RecipePreviewHandler, IInitializeWithFile),
        QITABENT(RecipePreviewHandler, IPreviewHandlerVisuals), 
        QITABENT(RecipePreviewHandler, IOleWindow), 
        QITABENT(RecipePreviewHandler, IObjectWithSite), 
        { 0 },
    };
    return QISearch(this, qit, riid, ppv);
}

// Increase the reference count for an interface on an object.
IFACEMETHODIMP_(ULONG) RecipePreviewHandler::AddRef()
{
    return InterlockedIncrement(&m_cRef);
}

// Decrease the reference count for an interface on an object.
IFACEMETHODIMP_(ULONG) RecipePreviewHandler::Release()
{
    ULONG cRef = InterlockedDecrement(&m_cRef);
    if (0 == cRef)
    {
        delete this;
    }

    return cRef;
}

#pragma endregion


#pragma region IInitializeWithFile

IFACEMETHODIMP RecipePreviewHandler::Initialize(LPCWSTR pszFilePath, DWORD grfMode)
{
    HRESULT hr = E_INVALIDARG;
    if (pszFilePath)
    {
      // Initialize can be called more than once, so release existing valid 
      // m_pStream.
        if (m_pPathFile)
        {
            LocalFree(m_pPathFile);
            m_pPathFile = NULL;
        }

        m_pPathFile = StrDup(pszFilePath);
        hr = S_OK;
    }
    return hr;
}
#pragma endregion


#pragma region IPreviewHandler

// This method gets called when the previewer gets created. It sets the parent 
// window of the previewer window, as well as the area within the parent to be 
// used for the previewer window.
IFACEMETHODIMP RecipePreviewHandler::SetWindow(HWND hwnd, const RECT *prc)
{
    if (hwnd && prc)
    {
        m_hwndParent = hwnd;  // Cache the HWND for later use
        m_rcParent = *prc;    // Cache the RECT for later use

        if (m_hwndPreview)
        {
            // Update preview window parent and rect information
            SetParent(m_hwndPreview, m_hwndParent);
            SetWindowPos(m_hwndPreview, NULL, m_rcParent.left, m_rcParent.top, 
                RECTWIDTH(m_rcParent), RECTHEIGHT(m_rcParent), 
                SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);
        }
    }
    return S_OK;
}

// Directs the preview handler to set focus to itself.
IFACEMETHODIMP RecipePreviewHandler::SetFocus()
{
    HRESULT hr = S_FALSE;
    if (m_hwndPreview)
    {
        ::SetFocus(m_hwndPreview);
        hr = S_OK;
    }
    return hr;
}

// Directs the preview handler to return the HWND from calling the GetFocus 
// function.
IFACEMETHODIMP RecipePreviewHandler::QueryFocus(HWND *phwnd)
{
    HRESULT hr = E_INVALIDARG;
    if (phwnd)
    {
        *phwnd = ::GetFocus();
        if (*phwnd)
        {
            hr = S_OK;
        }
        else
        {
            hr = HRESULT_FROM_WIN32(GetLastError());
        }
    }
    return hr;
}

// Directs the preview handler to handle a keystroke passed up from the 
// message pump of the process in which the preview handler is running.
HRESULT RecipePreviewHandler::TranslateAccelerator(MSG *pmsg)
{
    HRESULT hr = S_FALSE;
    IPreviewHandlerFrame *pFrame = NULL;
    if (m_punkSite && SUCCEEDED(m_punkSite->QueryInterface(&pFrame)))
    {
        // If your previewer has multiple tab stops, you will need to do 
        // appropriate first/last child checking. This sample previewer has 
        // no tabstops, so we want to just forward this message out.
        hr = pFrame->TranslateAccelerator(pmsg);

        pFrame->Release();
    }
    return hr;
}

// This method gets called when the size of the previewer window changes 
// (user resizes the Reading Pane). It directs the preview handler to change 
// the area within the parent hwnd that it draws into.
IFACEMETHODIMP RecipePreviewHandler::SetRect(const RECT *prc)
{
    HRESULT hr = E_INVALIDARG;
    if (prc != NULL)
    {
        m_rcParent = *prc;
        if (m_hwndPreview)
        {
            // Preview window is already created, so set its size and position.
            SetWindowPos(m_hwndPreview, NULL, m_rcParent.left, m_rcParent.top,
                (m_rcParent.right - m_rcParent.left), // Width
                (m_rcParent.bottom - m_rcParent.top), // Height
                SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);

            SetWindowPos(GetDlgItem(m_hwndPreview, IDC_LIST_DATA), NULL, m_rcParent.left, m_rcParent.top,
                         (m_rcParent.right - m_rcParent.left)- CorrectWidth, // Width
                         (m_rcParent.bottom - m_rcParent.top) - (CorrectHeight), // Height
                         SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);
        }
        hr = S_OK;
    }
    return hr;
}

namespace fs = arrow::fs;



#define Convert(Type, Array, Getter) {  std::wstringstream conv; conv << std::static_pointer_cast<Type>(Array)->Getter;  return _wcsdup(conv.str().c_str()); }

LPWSTR asStr(std::shared_ptr<arrow::Field>& field, std::shared_ptr<arrow::Array>& array, int pos)
{
    switch (field->type()->id())
    {
    case arrow::Type::type::UINT8:
        Convert(arrow::UInt8Array, array, Value(pos)); break;
    case arrow::Type::type::INT8:
        Convert(arrow::Int8Array, array, Value(pos)); break;
    case arrow::Type::type::UINT16:
        Convert(arrow::UInt16Array, array, Value(pos)); break;
    case arrow::Type::type::INT16:
        Convert(arrow::Int16Array, array, Value(pos)); break;
    case arrow::Type::type::UINT32:
        Convert(arrow::UInt32Array, array, Value(pos)); break;
    case arrow::Type::type::INT32:
        Convert(arrow::Int32Array, array, Value(pos)); break;
    case arrow::Type::type::UINT64:
        Convert(arrow::UInt64Array, array, Value(pos)); break;
    case arrow::Type::type::INT64:
        Convert(arrow::Int64Array, array, Value(pos)); break;
    case arrow::Type::type::HALF_FLOAT:
        Convert(arrow::HalfFloatArray, array, Value(pos)); break;
    case arrow::Type::type::FLOAT:
        Convert(arrow::FloatArray, array, Value(pos)); break;
    case arrow::Type::type::DOUBLE:
        Convert(arrow::DoubleArray, array, Value(pos)); break;

    case arrow::Type::type::STRING:
        Convert(arrow::StringArray, array, GetString(pos).c_str()); break;


    }
    return _wcsdup(L"-");// toStr<field->type()>(array, pos);
}
//C:\Users\NicolasWiestDaessle\Documents\CheckoutDB\DCM\Checkout_Results\BV
// The method directs the preview handler to load data from the source 
// specified in an earlier Initialize method call, and to begin rendering to 
// the previewer window.
IFACEMETHODIMP RecipePreviewHandler::DoPreview()
{

    std::cerr << "FTH Preview: DoPreview";
    // Cannot call more than once.
    // (Unload should be called before another DoPreview)
    if (m_hwndPreview != NULL || !m_pPathFile)
    {
        return E_FAIL;
    }

    HRESULT hr = E_FAIL;   
    
    std::string uri = CW2A(m_pPathFile), root_path;
    auto fs = fs::FileSystemFromUriOrPath(uri, &root_path);
    if (!fs.ok())     return E_FAIL;
    auto input = fs.ValueOrDie()->OpenInputFile(uri);
    if (!input.ok())  return E_FAIL;

    auto reader = arrow::ipc::RecordBatchFileReader::Open(input.ValueOrDie());
    if (!reader.ok()) return E_FAIL;

    m_hwndPreview = CreateDialog(g_hInst, MAKEINTRESOURCE(IDD_MAINDIALOG), m_hwndParent, NULL);

    if (!m_hwndPreview)
    {
        auto retval = GetLastError();
        std::cerr << "Error" << retval;
        return HRESULT_FROM_WIN32(retval);
    }

    int cx = RECTWIDTH(m_rcParent), cy = RECTHEIGHT(m_rcParent);

    SetWindowPos(m_hwndPreview, NULL, m_rcParent.left, m_rcParent.top,
                 cx, cy,
                 SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);

             // Set the title label on the window.
    HWND hLabelTitle = GetDlgItem(m_hwndPreview, IDC_STATIC_TITLE);
    Static_SetText(hLabelTitle, m_pPathFile);



    auto schema = reader.ValueOrDie()->schema();

    // Skip metadata for the moment
    HWND hlistView = GetDlgItem(m_hwndPreview, IDC_LIST_DATA);

    SetWindowPos(hlistView, NULL, m_rcParent.left, m_rcParent.top,
                 (m_rcParent.right - m_rcParent.left) - CorrectWidth, // Width
                 (m_rcParent.bottom - m_rcParent.top) - (CorrectHeight), // Height
                 SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE); 
    
    int icol = 0;
    LVCOLUMN lvc;
   
    for (auto f : schema->fields())
    {
        std::wstringstream conv;   conv << f->name().c_str();  stored.push_back(_wcsdup(conv.str().c_str()));


        lvc.mask = LVCF_FMT | LVCF_WIDTH | LVCF_TEXT | LVCF_SUBITEM;

        lvc.iSubItem = icol;
        lvc.pszText = stored.back(); // uname;
        lvc.cx = 100;
        lvc.fmt = LVCFMT_RIGHT;
        ListView_InsertColumn(hlistView, icol, &lvc);
        ListView_SetColumnWidth(hlistView, icol, LVSCW_AUTOSIZE_USEHEADER); 
        icol++;
    }

    LVITEM item;

    item.mask = LVIF_TEXT;
    
    for (int record = 0; record < reader.ValueOrDie()->num_record_batches(); ++record)
    {
        auto rb = reader.ValueOrDie()->ReadRecordBatch(record).ValueOrDie();

        int nItem = 0;
        for (auto field : schema->fields())
        {
            auto array = rb->GetColumnByName(field->name());

            for (int s = 0; s < rb->num_rows(); ++s)
            {                               
                stored.push_back(asStr(field, array, s));

                item.pszText = stored.back();
                item.iItem = s;
                item.iSubItem = nItem;
                if (nItem == 0)
                    ListView_InsertItem(hlistView, &item);
                else
                    ListView_SetItem(hlistView, &item);
            }
            nItem++;
        } 

   //     break; // only do one record
    }
    
    std::wstringstream cols, rows;

    cols << schema->num_fields();
    rows << reader.ValueOrDie()->CountRows().ValueOrDie();

    HWND colDlg = GetDlgItem(m_hwndPreview, IDC_STATIC_COLS);

    stored.push_back(_wcsdup(cols.str().c_str()));
    Static_SetText(colDlg, stored.back());

    HWND rowDlg = GetDlgItem(m_hwndPreview, IDC_STATIC_ROWS);
    
    stored.push_back(_wcsdup(rows.str().c_str()));
    Static_SetText(rowDlg, stored.back());
    
    ShowWindow(m_hwndPreview, SW_SHOW);
    return S_OK;
}

// This method gets called when a shell item is de-selected. It directs the 
// preview handler to cease rendering a preview and to release all resources 
// that have been allocated based on the item passed in during the 
// initialization.
HRESULT RecipePreviewHandler::Unload()
{
    if (m_pPathFile)
    {
        LocalFree(m_pPathFile);
        m_pPathFile = NULL;
    }

    if (m_hwndPreview)
    {
        DestroyWindow(m_hwndPreview);
        m_hwndPreview = NULL;
    }

    for (auto str : stored)
        free(str);
    stored.clear();

    return S_OK;
}

#pragma endregion


#pragma region IPreviewHandlerVisuals (Optional)

// Sets the background color of the preview handler.
IFACEMETHODIMP RecipePreviewHandler::SetBackgroundColor(COLORREF color)
{
    return S_OK;
}

// Sets the font attributes to be used for text within the preview handler.
IFACEMETHODIMP RecipePreviewHandler::SetFont(const LOGFONTW *plf)
{
    return S_OK;
}

// Sets the color of the text within the preview handler.
IFACEMETHODIMP RecipePreviewHandler::SetTextColor(COLORREF color)
{
    return S_OK;
}

#pragma endregion


#pragma region IOleWindow

// Retrieves a handle to one of the windows participating in in-place 
// activation (frame, document, parent, or in-place object window).
IFACEMETHODIMP RecipePreviewHandler::GetWindow(HWND *phwnd)
{
    HRESULT hr = E_INVALIDARG;
    if (phwnd)
    {
        *phwnd = m_hwndParent;
        hr = S_OK;
    }
    return hr;
}

// Determines whether context-sensitive help mode should be entered during an 
// in-place activation session
IFACEMETHODIMP RecipePreviewHandler::ContextSensitiveHelp(BOOL fEnterMode)
{
    return E_NOTIMPL;
}

#pragma endregion


#pragma region IObjectWithSite

// Provides the site's IUnknown pointer to the object.
IFACEMETHODIMP RecipePreviewHandler::SetSite(IUnknown *punkSite)
{
    if (m_punkSite)
    {
        m_punkSite->Release();
        m_punkSite = NULL;
    }
    return punkSite ? punkSite->QueryInterface(&m_punkSite) : S_OK;
}

// Gets the last site set with IObjectWithSite::SetSite. If there is no known 
// site, the object returns a failure code.
IFACEMETHODIMP RecipePreviewHandler::GetSite(REFIID riid, void **ppv)
{
    *ppv = NULL;
    return m_punkSite ? m_punkSite->QueryInterface(riid, ppv) : E_FAIL;
}

#pragma endregion

