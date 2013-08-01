#include "cppwebpage.h"
#include <stdio.h>

HRESULT STDMETHODCALLTYPE MyIOleInPlaceFrame::QueryInterface( 
    /* [in] */ REFIID riid,
    /* [iid_is][out] */ __RPC__deref_out void __RPC_FAR *__RPC_FAR *ppvObject)
{
	NOTIMPLEMENTED;
}

ULONG STDMETHODCALLTYPE MyIOleInPlaceFrame::AddRef( void)
{
	return(1);
}

ULONG STDMETHODCALLTYPE MyIOleInPlaceFrame::Release( void)
{
	return(1);
}

/* [input_sync] */ HRESULT STDMETHODCALLTYPE MyIOleInPlaceFrame::GetWindow( 
    /* [out] */ __RPC__deref_out_opt HWND *phwnd)
{
	*phwnd = ((_IOleInPlaceFrameEx *)this)->window;
	return(S_OK);
}
        
HRESULT STDMETHODCALLTYPE MyIOleInPlaceFrame::ContextSensitiveHelp( 
    /* [in] */ BOOL fEnterMode)
{
	NOTIMPLEMENTED;
}

/* [input_sync] */ HRESULT STDMETHODCALLTYPE MyIOleInPlaceFrame::GetBorder( 
    /* [out] */ __RPC__out LPRECT lprectBorder)
{
	NOTIMPLEMENTED;
}
        
/* [input_sync] */ HRESULT STDMETHODCALLTYPE MyIOleInPlaceFrame::RequestBorderSpace( 
    /* [unique][in] */ __RPC__in_opt LPCBORDERWIDTHS pborderwidths)
{
	NOTIMPLEMENTED;
}
        
/* [input_sync] */ HRESULT STDMETHODCALLTYPE MyIOleInPlaceFrame::SetBorderSpace( 
    /* [unique][in] */ __RPC__in_opt LPCBORDERWIDTHS pborderwidths)
{
	NOTIMPLEMENTED;
}
        
HRESULT STDMETHODCALLTYPE MyIOleInPlaceFrame::SetActiveObject( 
    /* [unique][in] */ __RPC__in_opt IOleInPlaceActiveObject *pActiveObject,
    /* [unique][string][in] */ __RPC__in_opt_string LPCOLESTR pszObjName)
{
	return(S_OK);
}

HRESULT STDMETHODCALLTYPE MyIOleInPlaceFrame::InsertMenus( 
    /* [in] */ __RPC__in HMENU hmenuShared,
    /* [out][in] */ __RPC__inout LPOLEMENUGROUPWIDTHS lpMenuWidths)
{
	NOTIMPLEMENTED;
}
        
/* [input_sync] */ HRESULT STDMETHODCALLTYPE MyIOleInPlaceFrame::SetMenu( 
    /* [in] */ __RPC__in HMENU hmenuShared,
    /* [in] */ __RPC__in HOLEMENU holemenu,
    /* [in] */ __RPC__in HWND hwndActiveObject)
{
	return (S_OK);
}
        
HRESULT STDMETHODCALLTYPE MyIOleInPlaceFrame::RemoveMenus( 
    /* [in] */ __RPC__in HMENU hmenuShared)
{
	NOTIMPLEMENTED;
}
        
/* [input_sync] */ HRESULT STDMETHODCALLTYPE MyIOleInPlaceFrame::SetStatusText( 
    /* [unique][in] */ __RPC__in_opt LPCOLESTR pszStatusText)
{
	return(S_OK);
}
        
HRESULT STDMETHODCALLTYPE MyIOleInPlaceFrame::EnableModeless( 
    /* [in] */ BOOL fEnable)
{
	return(S_OK);
}
        
HRESULT STDMETHODCALLTYPE MyIOleInPlaceFrame::TranslateAccelerator( 
    /* [in] */ __RPC__in LPMSG lpmsg,
    /* [in] */ WORD wID)
{
	NOTIMPLEMENTED;
}

HRESULT STDMETHODCALLTYPE MyIOleClientSite::QueryInterface( 
    /* [in] */ REFIID riid,
    /* [iid_is][out] */ __RPC__deref_out void __RPC_FAR *__RPC_FAR *ppvObject)
{
	if (!memcmp(&riid, &IID_IUnknown, sizeof(GUID)) || !memcmp(&riid, &IID_IOleClientSite, sizeof(GUID)))
		*ppvObject = &((_IOleClientSiteEx *)this)->client;

	// If the browser is asking us to match IID_IOleInPlaceSite, then it wants us to return a pointer to
	// our IOleInPlaceSite struct. Then the browser will use the VTable in that struct to call our
	// IOleInPlaceSite functions.  It will also pass this same pointer to all of our IOleInPlaceSite
	// functions (except for Site_QueryInterface, Site_AddRef, and Site_Release. Those will always get
	// the pointer to our _IOleClientSiteEx.
	//
	// Actually, we're going to lie to the browser. We're going to return our own _IOleInPlaceSiteEx
	// struct, and tell the browser that it's a IOleInPlaceSite struct. It's ok. The first thing in
	// our _IOleInPlaceSiteEx is an embedded IOleInPlaceSite, so the browser doesn't mind. We want the
	// browser to continue passing our _IOleInPlaceSiteEx pointer wherever it would normally pass a
	// IOleInPlaceSite pointer.
	else if (!memcmp(&riid, &IID_IOleInPlaceSite, sizeof(GUID)))
		*ppvObject = &((_IOleClientSiteEx *)this)->inplace;

	// If the browser is asking us to match IID_IDocHostUIHandler, then it wants us to return a pointer to
	// our IDocHostUIHandler struct. Then the browser will use the VTable in that struct to call our
	// IDocHostUIHandler functions.  It will also pass this same pointer to all of our IDocHostUIHandler
	// functions (except for Site_QueryInterface, Site_AddRef, and Site_Release. Those will always get
	// the pointer to our _IOleClientSiteEx.
	//
	// Actually, we're going to lie to the browser. We're going to return our own _IDocHostUIHandlerEx
	// struct, and tell the browser that it's a IDocHostUIHandler struct. It's ok. The first thing in
	// our _IDocHostUIHandlerEx is an embedded IDocHostUIHandler, so the browser doesn't mind. We want the
	// browser to continue passing our _IDocHostUIHandlerEx pointer wherever it would normally pass a
	// IDocHostUIHandler pointer. My, we're really playing dirty tricks on the browser here. heheh.
	else if (!memcmp(&riid, &IID_IDocHostUIHandler, sizeof(GUID)))
		*ppvObject = &((_IOleClientSiteEx *)this)->ui;

	// For other types of objects the browser wants, just report that we don't have any such objects.
	// NOTE: If you want to add additional functionality to your browser hosting, you may need to
	// provide some more objects here. You'll have to investigate what the browser is asking for
	// (ie, what REFIID it is passing).
	else
	{
		*ppvObject = 0;
		return(E_NOINTERFACE);
	}

	return(S_OK);
}

ULONG STDMETHODCALLTYPE MyIOleClientSite::AddRef( void)
{
	return(1);
}

ULONG STDMETHODCALLTYPE MyIOleClientSite::Release( void)
{
	return(1);
}

HRESULT STDMETHODCALLTYPE MyIOleClientSite::SaveObject(void)
{
	NOTIMPLEMENTED;
}
        
HRESULT STDMETHODCALLTYPE MyIOleClientSite::GetMoniker( 
    /* [in] */ DWORD dwAssign,
    /* [in] */ DWORD dwWhichMoniker,
    /* [out] */ __RPC__deref_out_opt IMoniker **ppmk)
{
	NOTIMPLEMENTED;
}
        
HRESULT STDMETHODCALLTYPE MyIOleClientSite::GetContainer( 
    /* [out] */ __RPC__deref_out_opt IOleContainer **ppContainer)
{
	*ppContainer = 0;

	return(E_NOINTERFACE);
}
        
HRESULT STDMETHODCALLTYPE MyIOleClientSite::ShowObject(void)
{
	return(NOERROR);
}
        
HRESULT STDMETHODCALLTYPE MyIOleClientSite::OnShowWindow( 
    /* [in] */ BOOL fShow)
{
	NOTIMPLEMENTED;
}
        
HRESULT STDMETHODCALLTYPE MyIOleClientSite::RequestNewObjectLayout(void)
{
	NOTIMPLEMENTED;
}

HRESULT STDMETHODCALLTYPE MyIDocHostUIHandler::QueryInterface( 
    /* [in] */ REFIID riid,
    /* [iid_is][out] */ __RPC__deref_out void __RPC_FAR *__RPC_FAR *ppvObject)
{
	return ((IOleClientSite *)((char *)this - sizeof(IOleClientSite) - sizeof(_IOleInPlaceSiteEx)))->QueryInterface(riid, ppvObject);
}

ULONG STDMETHODCALLTYPE MyIDocHostUIHandler::AddRef(void)
{
	return(1);
}

ULONG STDMETHODCALLTYPE MyIDocHostUIHandler::Release(void)
{
	return(1);
}


HRESULT STDMETHODCALLTYPE MyIDocHostUIHandler::ShowContextMenu( 
    /* [in] */ DWORD dwID,
    /* [in] */ POINT *ppt,
    /* [in] */ IUnknown *pcmdtReserved,
    /* [in] */ IDispatch *pdispReserved)
{
	return(S_OK);
}
        
HRESULT STDMETHODCALLTYPE MyIDocHostUIHandler::GetHostInfo( 
    /* [out][in] */ DOCHOSTUIINFO *pInfo)
{
	pInfo->cbSize = sizeof(DOCHOSTUIINFO);

	pInfo->dwFlags = DOCHOSTUIFLAG_NO3DBORDER;

	pInfo->dwDoubleClick = DOCHOSTUIDBLCLK_DEFAULT;

	return(S_OK);
}
        
HRESULT STDMETHODCALLTYPE MyIDocHostUIHandler::ShowUI( 
    /* [in] */ DWORD dwID,
    /* [in] */ IOleInPlaceActiveObject *pActiveObject,
    /* [in] */ IOleCommandTarget *pCommandTarget,
    /* [in] */ IOleInPlaceFrame *pFrame,
    /* [in] */ IOleInPlaceUIWindow *pDoc)
{
	return(S_OK);
}
        
HRESULT STDMETHODCALLTYPE MyIDocHostUIHandler::HideUI(void)
{
	return(S_OK);
}
        
HRESULT STDMETHODCALLTYPE MyIDocHostUIHandler::UpdateUI(void)
{
	return(S_OK);
}
        
HRESULT STDMETHODCALLTYPE MyIDocHostUIHandler::EnableModeless( 
    /* [in] */ BOOL fEnable)
{
	return(S_OK);
}
        
HRESULT STDMETHODCALLTYPE MyIDocHostUIHandler::OnDocWindowActivate( 
    /* [in] */ BOOL fActivate)
{
	return(S_OK);
}
        
HRESULT STDMETHODCALLTYPE MyIDocHostUIHandler::OnFrameWindowActivate( 
    /* [in] */ BOOL fActivate)
{
	return(S_OK);
}
        
HRESULT STDMETHODCALLTYPE MyIDocHostUIHandler::ResizeBorder( 
    /* [in] */ LPCRECT prcBorder,
    /* [in] */ IOleInPlaceUIWindow *pUIWindow,
    /* [in] */ BOOL fRameWindow)
{
	return(S_OK);
}
        
HRESULT STDMETHODCALLTYPE MyIDocHostUIHandler::TranslateAccelerator( 
    /* [in] */ LPMSG lpMsg,
    /* [in] */ const GUID *pguidCmdGroup,
    /* [in] */ DWORD nCmdID)
{
	return(S_FALSE);
}


HRESULT STDMETHODCALLTYPE MyIDocHostUIHandler::GetOptionKeyPath( 
    /* [annotation][out] */ 
    __out  LPOLESTR *pchKey,
    /* [in] */ DWORD dw)
{
	return(S_FALSE);
}
        
HRESULT STDMETHODCALLTYPE MyIDocHostUIHandler::GetDropTarget( 
    /* [in] */ IDropTarget *pDropTarget,
    /* [out] */ IDropTarget **ppDropTarget)
{
	return(S_FALSE);
}
        
HRESULT STDMETHODCALLTYPE MyIDocHostUIHandler::GetExternal( 
    /* [out] */ IDispatch **ppDispatch)
{
	*ppDispatch = 0;
	return(S_FALSE);
}
        
HRESULT STDMETHODCALLTYPE MyIDocHostUIHandler::TranslateUrl( 
    /* [in] */ DWORD dwTranslate,
    /* [annotation][in] */ 
    __in __nullterminated  OLECHAR *pchURLIn,
    /* [annotation][out] */ 
    __out  OLECHAR **ppchURLOut)
{
	*ppchURLOut = 0;
    return(S_FALSE);
}
        
HRESULT STDMETHODCALLTYPE MyIDocHostUIHandler::FilterDataObject( 
    /* [in] */ IDataObject *pDO,
    /* [out] */ IDataObject **ppDORet)
{
	*ppDORet = 0;
	return(S_FALSE);
}

HRESULT STDMETHODCALLTYPE MyIOleInPlaceSite::QueryInterface( 
    /* [in] */ REFIID riid,
    /* [iid_is][out] */ __RPC__deref_out void __RPC_FAR *__RPC_FAR *ppvObject)
{
	return (((MyIOleClientSite *)this) - 1)->QueryInterface(riid, ppvObject);
}

ULONG STDMETHODCALLTYPE MyIOleInPlaceSite::AddRef( void)
{
	return(1);
}

ULONG STDMETHODCALLTYPE MyIOleInPlaceSite::Release( void)
{
	return(1);
}

/* [input_sync] */ HRESULT STDMETHODCALLTYPE MyIOleInPlaceSite::GetWindow( 
    /* [out] */ __RPC__deref_out_opt HWND *phwnd)
{
	*phwnd = ((_IOleInPlaceSiteEx *)this)->frame.window;

	return(S_OK);
}
        
HRESULT STDMETHODCALLTYPE MyIOleInPlaceSite::ContextSensitiveHelp( 
    /* [in] */ BOOL fEnterMode)
{
	NOTIMPLEMENTED;
}

HRESULT STDMETHODCALLTYPE MyIOleInPlaceSite::CanInPlaceActivate(void)
{
	return(S_OK);
}
        
HRESULT STDMETHODCALLTYPE MyIOleInPlaceSite::OnInPlaceActivate(void)
{
	return(S_OK);
}
        
HRESULT STDMETHODCALLTYPE MyIOleInPlaceSite::OnUIActivate(void)
{
	return(S_OK);
}
        
HRESULT STDMETHODCALLTYPE MyIOleInPlaceSite::GetWindowContext( 
    /* [out] */ __RPC__deref_out_opt IOleInPlaceFrame **ppFrame,
    /* [out] */ __RPC__deref_out_opt IOleInPlaceUIWindow **ppDoc,
    /* [out] */ __RPC__out LPRECT lprcPosRect,
    /* [out] */ __RPC__out LPRECT lprcClipRect,
    /* [out][in] */ __RPC__inout LPOLEINPLACEFRAMEINFO lpFrameInfo)
{
	*ppFrame = (LPOLEINPLACEFRAME)&((_IOleInPlaceSiteEx *)this)->frame;

	// We have no OLEINPLACEUIWINDOW
	*ppDoc = 0;

	// Fill in some other info for the browser
	lpFrameInfo->fMDIApp = FALSE;
	lpFrameInfo->hwndFrame = ((_IOleInPlaceFrameEx *)*ppFrame)->window;
	lpFrameInfo->haccel = 0;
	lpFrameInfo->cAccelEntries = 0;

	// Give the browser the dimensions of where it can draw. We give it our entire window to fill.
	// We do this in InPlace_OnPosRectChange() which is called right when a window is first
	// created anyway, so no need to duplicate it here.
//	GetClientRect(lpFrameInfo->hwndFrame, lprcPosRect);
//	GetClientRect(lpFrameInfo->hwndFrame, lprcClipRect);

	return(S_OK);
}
        
HRESULT STDMETHODCALLTYPE MyIOleInPlaceSite::Scroll( 
    /* [in] */ SIZE scrollExtant)
{
	NOTIMPLEMENTED;
}
        
HRESULT STDMETHODCALLTYPE MyIOleInPlaceSite::OnUIDeactivate( 
    /* [in] */ BOOL fUndoable)
{
	return(S_OK);
}
        
HRESULT STDMETHODCALLTYPE MyIOleInPlaceSite::OnInPlaceDeactivate(void)
{
	return(S_OK);
}
        
HRESULT STDMETHODCALLTYPE MyIOleInPlaceSite::DiscardUndoState(void)
{
	return(S_OK);
}
        
HRESULT STDMETHODCALLTYPE MyIOleInPlaceSite::DeactivateAndUndo(void)
{
	NOTIMPLEMENTED;
}
        
HRESULT STDMETHODCALLTYPE MyIOleInPlaceSite::OnPosRectChange( 
    /* [in] */ __RPC__in LPCRECT lprcPosRect)
{
	IOleObject			*browserObject;
	IOleInPlaceObject	*inplace;

	// We need to get the browser's IOleInPlaceObject object so we can call its SetObjectRects
	// function.
	browserObject = *((IOleObject **)((char *)this - sizeof(IOleObject *) - sizeof(IOleClientSite)));
	if (!(browserObject->QueryInterface(IID_IOleInPlaceObject, (void**)&inplace)))
	{
		// Give the browser the dimensions of where it can draw.
		inplace->SetObjectRects(lprcPosRect, lprcPosRect);
	}

	return(S_OK);
}
























































































/*************************** UnEmbedBrowserObject() ************************
 * Called to detach the browser object from our host window, and free its
 * resources, right before we destroy our window.
 *
 * hwnd =		Handle to the window hosting the browser object.
 *
 * NOTE: The pointer to the browser object must have been stored in the
 * window's USERDATA field. In other words, don't call UnEmbedBrowserObject().
 * with a HWND that wasn't successfully passed to EmbedBrowserObject().
 */

void UnEmbedBrowserObject(HWND hwnd)
{
	IOleObject	**browserHandle;
	IOleObject	*browserObject;

	// Retrieve the browser object's pointer we stored in our window's GWL_USERDATA when we
	// initially attached the browser object to this window. If 0, you may have called this
	// for a window that wasn't successfully passed to EmbedBrowserObject(). Bad boy! Or, we
	// may have failed the EmbedBrowserObject() call in WM_CREATE, in which case, our window
	// may get a WM_DESTROY which could call this a second time (ie, since we may call
	// UnEmbedBrowserObject in EmbedBrowserObject).
	if ((browserHandle = (IOleObject **)GetWindowLong(hwnd, GWL_USERDATA)))
	{
		browserObject = *browserHandle;

		// Unembed the browser object, and release its resources.
		browserObject->Close(OLECLOSE_NOSAVE);
		browserObject->Release();

		// Zero out the pointer just in case UnEmbedBrowserObject is called again for this window
		SetWindowLong(hwnd, GWL_USERDATA, 0);
	}
}






#define WEBPAGE_GOBACK		0
#define WEBPAGE_GOFORWARD	1
#define WEBPAGE_GOHOME		2
#define WEBPAGE_SEARCH		3
#define WEBPAGE_REFRESH		4
#define WEBPAGE_STOP		5

/******************************* DoPageAction() **************************
 * Implements the functionality of a "Back". "Forward", "Home", "Search",
 * "Refresh", or "Stop" button.
 *
 * hwnd =		Handle to the window hosting the browser object.
 * action =		One of the following:
 *				0 = Move back to the previously viewed web page.
 *				1 = Move forward to the previously viewed web page.
 *				2 = Move to the home page.
 *				3 = Search.
 *				4 = Refresh the page.
 *				5 = Stop the currently loading page.
 *
 * NOTE: EmbedBrowserObject() must have been successfully called once with the
 * specified window, prior to calling this function. You need call
 * EmbedBrowserObject() once only, and then you can make multiple calls to
 * this function to display numerous pages in the specified window.
 */

void DoPageAction(HWND hwnd, DWORD action)
{	
	IWebBrowser2	*webBrowser2;
	IOleObject		*browserObject;

	// Retrieve the browser object's pointer we stored in our window's GWL_USERDATA when
	// we initially attached the browser object to this window.
	browserObject = *((IOleObject **)GetWindowLong(hwnd, GWL_USERDATA));

	// We want to get the base address (ie, a pointer) to the IWebBrowser2 object embedded within the browser
	// object, so we can call some of the functions in the former's table.
	if (!browserObject->QueryInterface(IID_IWebBrowser2, (void**)&webBrowser2))
	{
		// Ok, now the pointer to our IWebBrowser2 object is in 'webBrowser2', and so its VTable is
		// webBrowser2->lpVtbl.

		// Call the desired function
		switch (action)
		{
			case WEBPAGE_GOBACK:
			{
				// Call the IWebBrowser2 object's GoBack function.
				webBrowser2->GoBack();
				break;
			}

			case WEBPAGE_GOFORWARD:
			{
				// Call the IWebBrowser2 object's GoForward function.
				webBrowser2->GoForward();
				break;
			}

			case WEBPAGE_GOHOME:
			{
				// Call the IWebBrowser2 object's GoHome function.
				webBrowser2->GoHome();
				break;
			}

			case WEBPAGE_SEARCH:
			{
				// Call the IWebBrowser2 object's GoSearch function.
				webBrowser2->GoSearch();
				break;
			}

			case WEBPAGE_REFRESH:
			{
				// Call the IWebBrowser2 object's Refresh function.
				webBrowser2->Refresh();
			}

			case WEBPAGE_STOP:
			{
				// Call the IWebBrowser2 object's Stop function.
				webBrowser2->Stop();
			}
		}

		// Release the IWebBrowser2 object.
		webBrowser2->Release();
	}
}





/******************************* DisplayHTMLStr() ****************************
 * Takes a string containing some HTML BODY, and displays it in the specified
 * window. For example, perhaps you want to display the HTML text of...
 *
 * <P>This is a picture.<P><IMG src="mypic.jpg">
 *
 * hwnd =		Handle to the window hosting the browser object.
 * string =		Pointer to nul-terminated string containing the HTML BODY.
 *				(NOTE: No <BODY></BODY> tags are required in the string).
 *
 * RETURNS: 0 if success, or non-zero if an error.
 *
 * NOTE: EmbedBrowserObject() must have been successfully called once with the
 * specified window, prior to calling this function. You need call
 * EmbedBrowserObject() once only, and then you can make multiple calls to
 * this function to display numerous pages in the specified window.
 */

long DisplayHTMLStr(HWND hwnd, LPCTSTR string)
{	
	IWebBrowser2	*webBrowser2;
	LPDISPATCH		lpDispatch;
	IHTMLDocument2	*htmlDoc2;
	IOleObject		*browserObject;
	SAFEARRAY		*sfArray;
	VARIANT			myURL;
	VARIANT			*pVar;
	BSTR			bstr;

	// Retrieve the browser object's pointer we stored in our window's GWL_USERDATA when
	// we initially attached the browser object to this window.
	browserObject = *((IOleObject **)GetWindowLong(hwnd, GWL_USERDATA));

	// Assume an error.
	bstr = 0;

	// We want to get the base address (ie, a pointer) to the IWebBrowser2 object embedded within the browser
	// object, so we can call some of the functions in the former's table.
	if (!browserObject->QueryInterface(IID_IWebBrowser2, (void**)&webBrowser2))
	{
		// Ok, now the pointer to our IWebBrowser2 object is in 'webBrowser2', and so its VTable is
		// webBrowser2->lpVtbl.

		// Call the IWebBrowser2 object's get_Document so we can get its DISPATCH object. I don't know why you
		// don't get the DISPATCH object via the browser object's QueryInterface(), but you don't.
		lpDispatch = 0;
		webBrowser2->get_Document(&lpDispatch);
		if (!lpDispatch)
		{
			// Before we can get_Document(), we actually need to have some HTML page loaded in the browser. So,
			// let's load an empty HTML page. Then, once we have that empty page, we can get_Document() and
			// write() to stuff our HTML string into it.
			VariantInit(&myURL);
			myURL.vt = VT_BSTR;
			myURL.bstrVal = SysAllocString(&Blank[0]);

			// Call the Navigate2() function to actually display the page.
			webBrowser2->Navigate2(&myURL, 0, 0, 0, 0);

			// Free any resources (including the BSTR).
			VariantClear(&myURL);
		}

		// Call the IWebBrowser2 object's get_Document so we can get its DISPATCH object. I don't know why you
		// don't get the DISPATCH object via the browser object's QueryInterface(), but you don't.
		if (!webBrowser2->get_Document(&lpDispatch) && lpDispatch)
		{
			// We want to get a pointer to the IHTMLDocument2 object embedded within the DISPATCH
			// object, so we can call some of the functions in the former's table.
			if (!lpDispatch->QueryInterface(IID_IHTMLDocument2, (void**)&htmlDoc2))
			{
				// Ok, now the pointer to our IHTMLDocument2 object is in 'htmlDoc2', and so its VTable is
				// htmlDoc2->lpVtbl.

				// Our HTML must be in the form of a BSTR. And it must be passed to write() in an
				// array of "VARIENT" structs. So let's create all that.
				if ((sfArray = SafeArrayCreate(VT_VARIANT, 1, (SAFEARRAYBOUND *)&ArrayBound)))
				{
					if (!SafeArrayAccessData(sfArray, (void**)&pVar))
					{
						pVar->vt = VT_BSTR;
#ifndef UNICODE
						{
						wchar_t		*buffer;
						DWORD		size;

						size = MultiByteToWideChar(CP_ACP, 0, string, -1, 0, 0);
						if (!(buffer = (wchar_t *)GlobalAlloc(GMEM_FIXED, sizeof(wchar_t) * size))) goto bad;
						MultiByteToWideChar(CP_ACP, 0, string, -1, buffer, size);
						bstr = SysAllocString(buffer);
						GlobalFree(buffer);
						}
#else
						bstr = SysAllocString(string);
#endif
						// Store our BSTR pointer in the VARIENT.
						if ((pVar->bstrVal = bstr))
						{
							// Pass the VARIENT with its BSTR to write() in order to shove our desired HTML string
							// into the body of that empty page we created above.
							htmlDoc2->write(sfArray);

							// Close the document. If we don't do this, subsequent calls to DisplayHTMLStr
							// would append to the current contents of the page
							htmlDoc2->close();

							// Normally, we'd need to free our BSTR, but SafeArrayDestroy() does it for us
//							SysFreeString(bstr);
						}
					}

					// Free the array. This also frees the VARIENT that SafeArrayAccessData created for us,
					// and even frees the BSTR we allocated with SysAllocString
					SafeArrayDestroy(sfArray);
				}

				// Release the IHTMLDocument2 object.
bad:			htmlDoc2->Release();
			}

			// Release the DISPATCH object.
			lpDispatch->Release();
		}

		// Release the IWebBrowser2 object.
		webBrowser2->Release();
	}

	// No error?
	if (bstr) return(0);

	// An error
	return(-1);
}






/******************************* DisplayHTMLPage() ****************************
 * Displays a URL, or HTML file on disk.
 *
 * hwnd =		Handle to the window hosting the browser object.
 * webPageName =	Pointer to nul-terminated name of the URL/file.
 *
 * RETURNS: 0 if success, or non-zero if an error.
 *
 * NOTE: EmbedBrowserObject() must have been successfully called once with the
 * specified window, prior to calling this function. You need call
 * EmbedBrowserObject() once only, and then you can make multiple calls to
 * this function to display numerous pages in the specified window.
 */

long DisplayHTMLPage(HWND hwnd, LPTSTR webPageName)
{
	IWebBrowser2	*webBrowser2;
	VARIANT			myURL;
	IOleObject		*browserObject;

	// Retrieve the browser object's pointer we stored in our window's GWL_USERDATA when
	// we initially attached the browser object to this window.
	browserObject = *((IOleObject **)GetWindowLong(hwnd, GWL_USERDATA));

	// We want to get the base address (ie, a pointer) to the IWebBrowser2 object embedded within the browser
	// object, so we can call some of the functions in the former's table.
	if (!browserObject->QueryInterface(IID_IWebBrowser2, (void**)&webBrowser2))
	{
		// Ok, now the pointer to our IWebBrowser2 object is in 'webBrowser2', and so its VTable is
		// webBrowser2->lpVtbl.

		// Our URL (ie, web address, such as "http://www.microsoft.com" or an HTM filename on disk
		// such as "c:\myfile.htm") must be passed to the IWebBrowser2's Navigate2() function as a BSTR.
		// A BSTR is like a pascal version of a double-byte character string. In other words, the
		// first unsigned short is a count of how many characters are in the string, and then this
		// is followed by those characters, each expressed as an unsigned short (rather than a
		// char). The string is not nul-terminated. The OS function SysAllocString can allocate and
		// copy a UNICODE C string to a BSTR. Of course, we'll need to free that BSTR after we're done
		// with it. If we're not using UNICODE, we first have to convert to a UNICODE string.
		//
		// What's more, our BSTR needs to be stuffed into a VARIENT struct, and that VARIENT struct is
		// then passed to Navigate2(). Why? The VARIENT struct makes it possible to define generic
		// 'datatypes' that can be used with all languages. Not all languages support things like
		// nul-terminated C strings. So, by using a VARIENT, whose first field tells what sort of
		// data (ie, string, float, etc) is in the VARIENT, COM interfaces can be used by just about
		// any language.
		VariantInit(&myURL);
		myURL.vt = VT_BSTR;

#ifndef UNICODE
		{
		wchar_t		*buffer;
		DWORD		size;

		size = MultiByteToWideChar(CP_ACP, 0, webPageName, -1, 0, 0);
		if (!(buffer = (wchar_t *)GlobalAlloc(GMEM_FIXED, sizeof(wchar_t) * size))) goto badalloc;
		MultiByteToWideChar(CP_ACP, 0, webPageName, -1, buffer, size);
		myURL.bstrVal = SysAllocString(buffer);
		GlobalFree(buffer);
		}
#else
		myURL.bstrVal = SysAllocString(webPageName);
#endif
		if (!myURL.bstrVal)
		{
badalloc:	webBrowser2->Release();
			return(-6);
		}

		// Call the Navigate2() function to actually display the page.
		webBrowser2->Navigate2(&myURL, 0, 0, 0, 0);

		// Free any resources (including the BSTR we allocated above).
		VariantClear(&myURL);

		// We no longer need the IWebBrowser2 object (ie, we don't plan to call any more functions in it,
		// so we can release our hold on it). Note that we'll still maintain our hold on the browser
		// object.
		webBrowser2->Release();

		// Success
		return(0);
	}

	return(-5);
}





/******************************* ResizeBrowser() ****************************
 * Resizes the browser object for the specified window to the specified
 * width and height.
 *
 * hwnd =			Handle to the window hosting the browser object.
 * width =			Width.
 * height =			Height.
 *
 * NOTE: EmbedBrowserObject() must have been successfully called once with the
 * specified window, prior to calling this function. You need call
 * EmbedBrowserObject() once only, and then you can make multiple calls to
 * this function to resize the browser object.
 */

void ResizeBrowser(HWND hwnd, DWORD width, DWORD height)
{
	IWebBrowser2	*webBrowser2;
	IOleObject		*browserObject;

	// Retrieve the browser object's pointer we stored in our window's GWL_USERDATA when
	// we initially attached the browser object to this window.
	browserObject = *((IOleObject **)GetWindowLong(hwnd, GWL_USERDATA));

	// We want to get the base address (ie, a pointer) to the IWebBrowser2 object embedded within the browser
	// object, so we can call some of the functions in the former's table.
	if (!browserObject->QueryInterface(IID_IWebBrowser2, (void**)&webBrowser2))
	{
		// Ok, now the pointer to our IWebBrowser2 object is in 'webBrowser2', and so its VTable is
		// webBrowser2->lpVtbl.

		// Call are put_Width() and put_Height() to set the new width/height.
		webBrowser2->put_Width(width);
		webBrowser2->put_Height(height);

		// We no longer need the IWebBrowser2 object (ie, we don't plan to call any more functions in it,
		// so we can release our hold on it). Note that we'll still maintain our hold on the browser
		// object.
		webBrowser2->Release();
	}
}





/***************************** EmbedBrowserObject() **************************
 * Puts the browser object inside our host window, and save a pointer to this
 * window's browser object in the window's GWL_USERDATA field.
 *
 * hwnd =		Handle of our window into which we embed the browser object.
 *
 * RETURNS: 0 if success, or non-zero if an error.
 *
 * NOTE: We tell the browser object to occupy the entire client area of the
 * window.
 *
 * NOTE: No HTML page will be displayed here. We can do that with a subsequent
 * call to either DisplayHTMLPage() or DisplayHTMLStr(). This is merely once-only
 * initialization for using the browser object. In a nutshell, what we do
 * here is get a pointer to the browser object in our window's GWL_USERDATA
 * so we can access that object's functions whenever we want, and we also pass
 * pointers to our IOleClientSite, IOleInPlaceFrame, and IStorage structs so that
 * the browser can call our functions in our struct's VTables.
 */

long EmbedBrowserObject(HWND hwnd)
{
	LPCLASSFACTORY		pClassFactory;
	IOleObject			*browserObject;
	IWebBrowser2		*webBrowser2;
	RECT				rect;
	_IOleClientSiteEx	*_iOleClientSiteEx;

	// Our IOleClientSite, IOleInPlaceSite, and IOleInPlaceFrame functions need to get our window handle. We
	// could store that in some global. But then, that would mean that our functions would work with only that
	// one window. If we want to create multiple windows, each hosting its own browser object (to display its
	// own web page), then we need to create unique IOleClientSite, IOleInPlaceSite, and IOleInPlaceFrame
	// structs for each window. And we'll put an extra field at the end of those structs to store our extra
	// data such as a window handle. So, our functions won't have to touch global data, and can therefore be
	// re-entrant and work with multiple objects/windows.
	//
	// Remember that a pointer to our IOleClientSite we create here will be passed as the first arg to every
	// one of our IOleClientSite functions. Ditto with the IOleInPlaceFrame object we create here, and the
	// IOleInPlaceFrame functions. So, our functions are able to retrieve the window handle we'll store here,
	// and then, they'll work with all such windows containing a browser control.
	//
	// Furthermore, since the browser will be calling our IOleClientSite's QueryInterface to get a pointer to
	// our IOleInPlaceSite and IDocHostUIHandler objects, that means that our IOleClientSite QueryInterface
	// must have an easy way to grab those pointers. Probably the easiest thing to do is just embed our
	// IOleInPlaceSite and IDocHostUIHandler objects inside of an extended IOleClientSite which we'll call
	// a _IOleClientSiteEx. As long as they come after the pointer to the IOleClientSite VTable, then we're
	// ok.
	//
	// Of course, we need to GlobalAlloc the above structs now. We'll just get all 4 with a single call to
	// GlobalAlloc, especially since 3 of them are all contained inside of our _IOleClientSiteEx anyway.
	//
	// So, we're not actually allocating separate IOleClientSite, IOleInPlaceSite, IOleInPlaceFrame and
	// IDocHostUIHandler structs.
	//
	// One final thing. We're going to allocate extra room to store the pointer to the browser object.
	_ComIOleObjClientSiteEx *_ptr = new _ComIOleObjClientSiteEx();
	_ptr->iOleClientSiteExObj.inplace.frame.window = hwnd;
	_iOleClientSiteEx = (_IOleClientSiteEx *)((char *)_ptr + sizeof(IOleObject *));

	// Get a pointer to the browser object and lock it down (so it doesn't "disappear" while we're using
	// it in this program). We do this by calling the OS function OleCreate().
	//	
	// NOTE: We need this pointer to interact with and control the browser. With normal WIN32 controls such as a
	// Static, Edit, Combobox, etc, you obtain an HWND and send messages to it with SendMessage(). Not so with
	// the browser object. You need to get a pointer to its "base structure" (as returned by OleCreate()). This
	// structure contains an array of pointers to functions you can call within the browser object. Actually, the
	// base structure contains a 'lpVtbl' field that is a pointer to that array. We'll call the array a 'VTable'.
	//
	// For example, the browser object happens to have a SetHostNames() function we want to call. So, after we
	// retrieve the pointer to the browser object (in a local we'll name 'browserObject'), then we can call that
	// function, and pass it args, as so:
	//
	// browserObject->lpVtbl->SetHostNames(browserObject, SomeString, SomeString);
	//
	// There's our pointer to the browser object in 'browserObject'. And there's the pointer to the browser object's
	// VTable in 'browserObject->lpVtbl'. And the pointer to the SetHostNames function happens to be stored in an
	// field named 'SetHostNames' within the VTable. So we are actually indirectly calling SetHostNames by using
	// a pointer to it. That's how you use a VTable.
	//
	// NOTE: We pass our _IOleClientSiteEx struct and lie -- saying that it's a IOleClientSite. It's ok. A
	// _IOleClientSiteEx struct starts with an embedded IOleClientSite. So the browser won't care, and we want
	// this extended struct passed to our IOleClientSite functions.

	// Get a pointer to the browser object's IClassFactory object via CoGetClassObject()
	pClassFactory = 0;
	if (!CoGetClassObject(CLSID_WebBrowser, 3, NULL, IID_IClassFactory, (void **)&pClassFactory) && pClassFactory)
	{
		// Call the IClassFactory's CreateInstance() to create a browser object
		if (!pClassFactory->CreateInstance(0, IID_IOleObject, (void **)&browserObject))
		{
			// Free the IClassFactory. We need it only to create a browser object instance
			pClassFactory->Release();

			// Ok, we now have the pointer to the browser object in 'browserObject'. Let's save this in the
			// memory block we allocated above, and then save the pointer to that whole thing in our window's
			// USERDATA field. That way, if we need multiple windows each hosting its own browser object, we can
			// call EmbedBrowserObject() for each one, and easily associate the appropriate browser object with
			// its matching window and its own objects containing per-window data.
			*((IOleObject **)&(_ptr->iOleObj)) = browserObject;
			SetWindowLong(hwnd, GWL_USERDATA, (LONG)(_ptr));

			// Give the browser a pointer to my IOleClientSite object
			if (!browserObject->SetClientSite((IOleClientSite *)_iOleClientSiteEx))
			{
				// We can now call the browser object's SetHostNames function. SetHostNames lets the browser object know our
				// application's name and the name of the document in which we're embedding the browser. (Since we have no
				// document name, we'll pass a 0 for the latter). When the browser object is opened for editing, it displays
				// these names in its titlebar.
				//	
				// We are passing 3 args to SetHostNames. You'll note that the first arg to SetHostNames is the base
				// address of our browser control. This is something that you always have to remember when working in C
				// (as opposed to C++). When calling a VTable function, the first arg to that function must always be the
				// structure which contains the VTable. (In this case, that's the browser control itself). Why? That's
				// because that function is always assumed to be written in C++. And the first argument to any C++ function
				// must be its 'this' pointer (ie, the base address of its class, which in this case is our browser object
				// pointer). In C++, you don't have to pass this first arg, because the C++ compiler is smart enough to
				// produce an executable that always adds this first arg. In fact, the C++ compiler is smart enough to
				// know to fetch the function pointer from the VTable, so you don't even need to reference that. In other
				// words, the C++ equivalent code would be:
				//
				// browserObject->SetHostNames(L"My Host Name", 0);
				//
				// So, when you're trying to convert C++ code to C, always remember to add this first arg whenever you're
				// dealing with a VTable (ie, the field is usually named 'lpVtbl') in the standard objects, and also add
				// the reference to the VTable itself.
				//
				// Oh yeah, the L is because we need UNICODE strings. And BTW, the host and document names can be anything
				// you want.
				browserObject->SetHostNames(L"My Host Name", 0);

				GetClientRect(hwnd, &rect);

				// Let browser object know that it is embedded in an OLE container.
				if (!OleSetContainedObject((struct IUnknown *)browserObject, TRUE) &&

					// Set the display area of our browser control the same as our window's size
					// and actually put the browser object into our window.
					!browserObject->DoVerb(OLEIVERB_SHOW, NULL, (IOleClientSite *)_iOleClientSiteEx, -1, hwnd, &rect) &&

					// Ok, now things may seem to get even trickier, One of those function pointers in the browser's VTable is
					// to the QueryInterface() function. What does this function do? It lets us grab the base address of any
					// other object that may be embedded within the browser object. And this other object has its own VTable
					// containing pointers to more functions we can call for that object.
					//
					// We want to get the base address (ie, a pointer) to the IWebBrowser2 object embedded within the browser
					// object, so we can call some of the functions in the former's table. For example, one IWebBrowser2 function
					// we intend to call below will be Navigate2(). So we call the browser object's QueryInterface to get our
					// pointer to the IWebBrowser2 object.
					!browserObject->QueryInterface(IID_IWebBrowser2, (void**)&webBrowser2))
				{
					// Ok, now the pointer to our IWebBrowser2 object is in 'webBrowser2', and so its VTable is
					// webBrowser2->lpVtbl.

					// Let's call several functions in the IWebBrowser2 object to position the browser display area
					// in our window. The functions we call are put_Left(), put_Top(), put_Width(), and put_Height().
					// Note that we reference the IWebBrowser2 object's VTable to get pointers to those functions. And
					// also note that the first arg we pass to each is the pointer to the IWebBrowser2 object.
					webBrowser2->put_Left(0);
					webBrowser2->put_Top(0);
					webBrowser2->put_Width(rect.right);
					webBrowser2->put_Height(rect.bottom);

					// We no longer need the IWebBrowser2 object (ie, we don't plan to call any more functions in it
					// right now, so we can release our hold on it). Note that we'll still maintain our hold on the
					// browser object until we're done with that object.
					webBrowser2->Release();

					// Success
					return(0);
				}
			}

			// Something went wrong setting up the browser!
			UnEmbedBrowserObject(hwnd);
			return(-4);
		}

		pClassFactory->Release();
		delete _ptr;

		// Can't create an instance of the browser!
		return(-3);
	}

	delete _ptr;

	// Can't get the web browser's IClassFactory!
	return(-2);
}



