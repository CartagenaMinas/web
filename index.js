var oDmApp= null;
var oScript = null;
function AutoConnect()
    {
        oScript = new ActiveXObject("StudioCommon.ScriptHelper");
        oScript.initialize(window);
        oDmApp = oScript.getApplication();
     }

function Cubagem()
    {
        
        //seleccionar los bloques del modelo dentro de un perimetro

        oDmApp.ParseCommand("selexy &IN="+tbMB.value+" &PERIM="+tbPoligono.value+" &OUT=__temp1 *X=XC *Y=YC @OUTSIDE=0 @CHECKROT=1 @PRINT=0");

        //

        if (rdAU.checked){oDmApp.ParseCommand("tongrad &IN=__temp1 &OUT=__temp2 *DENSITY=DENSITY *F1=AU @FACTOR=1 @TRENAME=0 @SETABSNT=1 @DENSITY=2.8 @COLUMN=0 @ROW=0 @BENCH=0 @KEYTOL=0.00001 @EXCEL=0");}

        else{oDmApp.ParseCommand("tongrad &IN=__temp1 &OUT=__temp2 *DENSITY=DENSITY *F1=AU @FACTOR=1 @TRENAME=0 @SETABSNT=1 @DENSITY=2.8 @COLUMN=0 @ROW=0 @BENCH=0 @KEYTOL=0.00001 @EXCEL=0");}

        var oDmFile = new ActiveXObject("DmFile.DmTableADO");
        oDmFile.Open(oDmApp.ActiveProject.Folder + "\\__temp2.dm",true);
        oDmFile.MoveTo(1);

        var volume=oDmFile.GetNamedColumn("VOLUME")
        var massa=oDmFile.GetNamedColumn("TONNES")
        
        if (rdAU.checked) {var teor=oDmFile.GetNamedColumn("AU")}
        else {var teor=oDmFile.GetNamedColumn("CU")}
        
        oDmFile.Close()

        tbVol.value=volume
        tbMassa.value=massa
        tbTeor.value=teor

        oDmApp.ParseCommand("delete-file '__temp1' 'yes'")

        alert("Proceso Finalisado");

        //atribuir los resultados

    }

    function DisplayBrowser() 
    {
        oDmBrowser = oDmApp.ActiveProject.Browser;
        oDmBrowser.TypeFilter = oScript.DmFileType.dmNothing;
        oDmBrowser.Show(false);
        return oDmBrowser.FileName;
    }




function btnBrowse1_onclick() 
    {
        tbMB.value = DisplayBrowser();
        oScript.makeFieldsPicklist(tbMB);
    }

function btnBrowse2_onclick() 
    {
        tbPoligono.value = DisplayBrowser();
        oScript.makeFieldsPicklist(tbPoligono);
    }
