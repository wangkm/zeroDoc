var maxTime = 30    // in seconds
var sleepTime = 250 // in milliseconds

var objArgs, ifname, fso, PDFCreator, DefaultPrinter, ReadyState,
	i, c, Scriptname;

fso = new ActiveXObject("Scripting.FileSystemObject");

Scriptname = fso.GetFileName(WScript.ScriptFullname);

if (WScript.Version < 5.1){
	WScript.Quit();
}

if (WScript.arguments.length == 0){
	WScript.Quit();
}

PDFCreator = WScript.CreateObject("PDFCreator.clsPDFCreator", "PDFCreator_");
PDFCreator.cStart("/NoProcessingAtStartup");

PDFCreator.cOption("UseAutosave") = 1;
PDFCreator.cOption("UseAutosaveDirectory") = 1;
PDFCreator.cOption("AutosaveFormat") = 0;		// 0 = PDF

// store the old default printer
DefaultPrinter = PDFCreator.cDefaultprinter;
PDFCreator.cDefaultprinter = "PDFCreator";
PDFCreator.cClearcache();

for (i = 0; i < WScript.arguments.length; i++){
	ifname = WScript.arguments.item(i)

	if (!fso.FileExists(ifname)){
		break;
	}
	if (!PDFCreator.cIsPrintable(ifname)){
		WScript.Quit();
	}

	ReadyState = 0

	PDFCreator.cOption("AutosaveDirectory") = fso.GetParentFolderName(ifname);
	PDFCreator.cOption("AutosaveFilename") = fso.GetBaseName(ifname);
	PDFCreator.cPrintfile(ifname);
	PDFCreator.cPrinterStop = false;

	c = 0
	while ((ReadyState == 0) && (c < (maxTime * 1000 / sleepTime))){
 		c = c + 1;
 		WScript.Sleep(sleepTime);
	}
	if (ReadyState == 0){
		break;
	}
}

// restore the old default printer
PDFCreator.cDefaultprinter = DefaultPrinter

PDFCreator.cClearcache();
WScript.Sleep(200);
PDFCreator.cClose();

//--- PDFCreator events ---

function PDFCreator_eReady()
{
	ReadyState = 1
}

function PDFCreator_eError()
{
	WScript.Quit();
}