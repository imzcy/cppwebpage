#ifndef CPPWEBPAGE_H
#define CPPWEBPAGE_H

#include <windows.h>
#include <exdisp.h>		// Defines of stuff like IWebBrowser2. This is an include file with Visual C 6 and above
#include <mshtml.h>		// Defines of stuff like IHTMLDocument2. This is an include file with Visual C 6 and above
#include <mshtmhst.h>	// Defines of stuff like IDocHostUIHandler. This is an include file with Visual C 6 and above
#include <crtdbg.h>		// for _ASSERT()
#include <tchar.h>		// for _tcsnicmp which compares two UNICODE strings case-insensitive

#define NOTIMPLEMENTED _ASSERT(0); return(E_NOTIMPL)

static wchar_t Blank[] = {L"about:blank"};
static const SAFEARRAYBOUND ArrayBound = {1, 0};

class MyIOleInPlaceFrame : public IOleInPlaceFrame
{
public:
    HRESULT STDMETHODCALLTYPE QueryInterface( 
        /* [in] */ REFIID riid,
        /* [iid_is][out] */ __RPC__deref_out void __RPC_FAR *__RPC_FAR *ppvObject);

    ULONG STDMETHODCALLTYPE AddRef(void);

    ULONG STDMETHODCALLTYPE Release(void);

public: /* IOleWindow */
    /* [input_sync] */ HRESULT STDMETHODCALLTYPE GetWindow( 
        /* [out] */ __RPC__deref_out_opt HWND *phwnd);
        
    HRESULT STDMETHODCALLTYPE ContextSensitiveHelp( 
        /* [in] */ BOOL fEnterMode);

public: /* IDispatch */
    /* [input_sync] */ HRESULT STDMETHODCALLTYPE GetBorder( 
        /* [out] */ __RPC__out LPRECT lprectBorder);
        
    /* [input_sync] */ HRESULT STDMETHODCALLTYPE RequestBorderSpace( 
        /* [unique][in] */ __RPC__in_opt LPCBORDERWIDTHS pborderwidths);
        
    /* [input_sync] */ HRESULT STDMETHODCALLTYPE SetBorderSpace( 
        /* [unique][in] */ __RPC__in_opt LPCBORDERWIDTHS pborderwidths);
        
    HRESULT STDMETHODCALLTYPE SetActiveObject( 
        /* [unique][in] */ __RPC__in_opt IOleInPlaceActiveObject *pActiveObject,
        /* [unique][string][in] */ __RPC__in_opt_string LPCOLESTR pszObjName);


public: /* IOleInPlaceFrame */
    HRESULT STDMETHODCALLTYPE InsertMenus( 
        /* [in] */ __RPC__in HMENU hmenuShared,
        /* [out][in] */ __RPC__inout LPOLEMENUGROUPWIDTHS lpMenuWidths);
        
    /* [input_sync] */ HRESULT STDMETHODCALLTYPE SetMenu( 
        /* [in] */ __RPC__in HMENU hmenuShared,
        /* [in] */ __RPC__in HOLEMENU holemenu,
        /* [in] */ __RPC__in HWND hwndActiveObject);
        
    HRESULT STDMETHODCALLTYPE RemoveMenus( 
        /* [in] */ __RPC__in HMENU hmenuShared);
        
    /* [input_sync] */ HRESULT STDMETHODCALLTYPE SetStatusText( 
        /* [unique][in] */ __RPC__in_opt LPCOLESTR pszStatusText);
        
    HRESULT STDMETHODCALLTYPE EnableModeless( 
        /* [in] */ BOOL fEnable);
        
    HRESULT STDMETHODCALLTYPE TranslateAccelerator( 
        /* [in] */ __RPC__in LPMSG lpmsg,
        /* [in] */ WORD wID);
};

class MyIOleClientSite : public IOleClientSite
{
public:
    HRESULT STDMETHODCALLTYPE QueryInterface( 
        /* [in] */ REFIID riid,
        /* [iid_is][out] */ __RPC__deref_out void __RPC_FAR *__RPC_FAR *ppvObject);

    ULONG STDMETHODCALLTYPE AddRef(void);

    ULONG STDMETHODCALLTYPE Release(void);

public:
    HRESULT STDMETHODCALLTYPE SaveObject(void);
        
    HRESULT STDMETHODCALLTYPE GetMoniker( 
        /* [in] */ DWORD dwAssign,
        /* [in] */ DWORD dwWhichMoniker,
        /* [out] */ __RPC__deref_out_opt IMoniker **ppmk);
        
    HRESULT STDMETHODCALLTYPE GetContainer( 
        /* [out] */ __RPC__deref_out_opt IOleContainer **ppContainer);
        
    HRESULT STDMETHODCALLTYPE ShowObject(void);
        
    HRESULT STDMETHODCALLTYPE OnShowWindow( 
        /* [in] */ BOOL fShow);
        
    HRESULT STDMETHODCALLTYPE RequestNewObjectLayout(void);

};

class MyIDocHostUIHandler : public IDocHostUIHandler
{
public:
    HRESULT STDMETHODCALLTYPE QueryInterface( 
        /* [in] */ REFIID riid,
        /* [iid_is][out] */ __RPC__deref_out void __RPC_FAR *__RPC_FAR *ppvObject);

    ULONG STDMETHODCALLTYPE AddRef(void);

    ULONG STDMETHODCALLTYPE Release(void);

public:
    HRESULT STDMETHODCALLTYPE ShowContextMenu( 
        /* [in] */ DWORD dwID,
        /* [in] */ POINT *ppt,
        /* [in] */ IUnknown *pcmdtReserved,
        /* [in] */ IDispatch *pdispReserved);
        
    HRESULT STDMETHODCALLTYPE GetHostInfo( 
        /* [out][in] */ DOCHOSTUIINFO *pInfo);
        
    HRESULT STDMETHODCALLTYPE ShowUI( 
        /* [in] */ DWORD dwID,
        /* [in] */ IOleInPlaceActiveObject *pActiveObject,
        /* [in] */ IOleCommandTarget *pCommandTarget,
        /* [in] */ IOleInPlaceFrame *pFrame,
        /* [in] */ IOleInPlaceUIWindow *pDoc);
        
    HRESULT STDMETHODCALLTYPE HideUI(void);
        
    HRESULT STDMETHODCALLTYPE UpdateUI(void);
        
    HRESULT STDMETHODCALLTYPE EnableModeless( 
        /* [in] */ BOOL fEnable);
        
    HRESULT STDMETHODCALLTYPE OnDocWindowActivate( 
        /* [in] */ BOOL fActivate);
        
    HRESULT STDMETHODCALLTYPE OnFrameWindowActivate( 
        /* [in] */ BOOL fActivate);
        
    HRESULT STDMETHODCALLTYPE ResizeBorder( 
        /* [in] */ LPCRECT prcBorder,
        /* [in] */ IOleInPlaceUIWindow *pUIWindow,
        /* [in] */ BOOL fRameWindow);
        
    HRESULT STDMETHODCALLTYPE TranslateAccelerator( 
        /* [in] */ LPMSG lpMsg,
        /* [in] */ const GUID *pguidCmdGroup,
        /* [in] */ DWORD nCmdID);
        
    HRESULT STDMETHODCALLTYPE GetOptionKeyPath( 
        /* [annotation][out] */ 
        __out  LPOLESTR *pchKey,
        /* [in] */ DWORD dw);
        
    HRESULT STDMETHODCALLTYPE GetDropTarget( 
        /* [in] */ IDropTarget *pDropTarget,
        /* [out] */ IDropTarget **ppDropTarget);
        
    HRESULT STDMETHODCALLTYPE GetExternal( 
        /* [out] */ IDispatch **ppDispatch);
        
    HRESULT STDMETHODCALLTYPE TranslateUrl( 
        /* [in] */ DWORD dwTranslate,
        /* [annotation][in] */ 
        __in __nullterminated  OLECHAR *pchURLIn,
        /* [annotation][out] */ 
        __out  OLECHAR **ppchURLOut);
        
    HRESULT STDMETHODCALLTYPE FilterDataObject( 
        /* [in] */ IDataObject *pDO,
        /* [out] */ IDataObject **ppDORet);
        
};

class MyIOleInPlaceSite : public IOleInPlaceSite
{
public:
    HRESULT STDMETHODCALLTYPE QueryInterface( 
        /* [in] */ REFIID riid,
        /* [iid_is][out] */ __RPC__deref_out void __RPC_FAR *__RPC_FAR *ppvObject);

    ULONG STDMETHODCALLTYPE AddRef(void);

    ULONG STDMETHODCALLTYPE Release(void);

public: /* IOleWindow */
    /* [input_sync] */ HRESULT STDMETHODCALLTYPE GetWindow( 
        /* [out] */ __RPC__deref_out_opt HWND *phwnd);
        
    HRESULT STDMETHODCALLTYPE ContextSensitiveHelp( 
        /* [in] */ BOOL fEnterMode);

public:
    HRESULT STDMETHODCALLTYPE CanInPlaceActivate(void);
        
    HRESULT STDMETHODCALLTYPE OnInPlaceActivate(void);
        
    HRESULT STDMETHODCALLTYPE OnUIActivate(void);
        
    HRESULT STDMETHODCALLTYPE GetWindowContext( 
        /* [out] */ __RPC__deref_out_opt IOleInPlaceFrame **ppFrame,
        /* [out] */ __RPC__deref_out_opt IOleInPlaceUIWindow **ppDoc,
        /* [out] */ __RPC__out LPRECT lprcPosRect,
        /* [out] */ __RPC__out LPRECT lprcClipRect,
        /* [out][in] */ __RPC__inout LPOLEINPLACEFRAMEINFO lpFrameInfo);
        
    HRESULT STDMETHODCALLTYPE Scroll( 
        /* [in] */ SIZE scrollExtant);
        
    HRESULT STDMETHODCALLTYPE OnUIDeactivate( 
        /* [in] */ BOOL fUndoable);
        
    HRESULT STDMETHODCALLTYPE OnInPlaceDeactivate(void);
        
    HRESULT STDMETHODCALLTYPE DiscardUndoState(void);
        
    HRESULT STDMETHODCALLTYPE DeactivateAndUndo(void);
        
    HRESULT STDMETHODCALLTYPE OnPosRectChange( 
        /* [in] */ __RPC__in LPCRECT lprcPosRect);

};


typedef struct {
	MyIOleInPlaceFrame	frame;		// The IOleInPlaceFrame must be first!

	///////////////////////////////////////////////////
	// Here you add any extra variables that you need
	// to access in your IOleInPlaceFrame functions.
	// You don't want those functions to access global
	// variables, because then you couldn't use more
	// than one browser object. (ie, You couldn't have
	// multiple windows, each with its own embedded
	// browser object to display a different web page).
	//
	// So here is where I added my extra HWND that my
	// IOleInPlaceFrame function Frame_GetWindow() needs
	// to access.
	///////////////////////////////////////////////////
	HWND				window;
} _IOleInPlaceFrameEx;

typedef struct {
	MyIOleInPlaceSite			inplace;	// My IOleInPlaceSite object. Must be first with in _IOleInPlaceSiteEx.

	///////////////////////////////////////////////////
	// Here you add any extra variables that you need
	// to access in your IOleInPlaceSite functions.
	//
	// So here is where I added my IOleInPlaceFrame
	// struct. If you need extra variables, add them
	// at the end.
	///////////////////////////////////////////////////
	_IOleInPlaceFrameEx		frame;		// My IOleInPlaceFrame object. Must be first within my _IOleInPlaceFrameEx
} _IOleInPlaceSiteEx;

typedef struct {
	MyIDocHostUIHandler		ui;			// My IDocHostUIHandler object. Must be first.

	///////////////////////////////////////////////////
	// Here you add any extra variables that you need
	// to access in your IDocHostUIHandler functions.
	///////////////////////////////////////////////////
} _IDocHostUIHandlerEx;

typedef struct {
	MyIOleClientSite		client;		// My IOleClientSite object. Must be first.
	_IOleInPlaceSiteEx		inplace;	// My IOleInPlaceSite object. A convenient place to put it.
	_IDocHostUIHandlerEx	ui;			// My IDocHostUIHandler object. Must be first within my _IDocHostUIHandlerEx.

	///////////////////////////////////////////////////
	// Here you add any extra variables that you need
	// to access in your IOleClientSite functions.
	///////////////////////////////////////////////////
} _IOleClientSiteEx;

typedef struct {
	IOleObject				*iOleObj;
	_IOleClientSiteEx		iOleClientSiteExObj;
} _ComIOleObjClientSiteEx;









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

void UnEmbedBrowserObject(HWND hwnd);





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

void DoPageAction(HWND hwnd, DWORD action);


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

long DisplayHTMLStr(HWND hwnd, LPCTSTR string);


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

long DisplayHTMLPage(HWND hwnd, LPTSTR webPageName);


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

void ResizeBrowser(HWND hwnd, DWORD width, DWORD height);


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

long EmbedBrowserObject(HWND hwnd);

#endif // CPPWEBPAGE_H