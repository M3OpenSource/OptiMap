import System;
import System.Web;
import System.Windows;
import System.Collections;
import System.Windows.Controls;
import System.Windows.Forms;
import System.Windows.Forms.Integration;
import System.Windows.Media;
import System.Windows.Input;
import System.Windows.Forms;
import System.Drawing.Color;
import System.Management;
import System.Windows.Controls.Primitives;
import System.Media;
import MForms;
import Mango.UI;
import Mango.UI.Controls;
import Mango.Services;
import Mango.UI.Core;
import Mango.UI.Controls;
import Mango.UI.Core.Util;
import Mango.UI.Services;
import Mango.DesignSystem;
import Lawson.M3.MI;
import Mango.UI.Services.Lists;
import System.ComponentModel;
import System.Net.Mail;
import System.ComponentModel;
import System.Reflection;
import System.Data.Odbc;

package MForms.JScript {
    class Optimap_v4 {

    var debug,content,controller;
    var addressFields: String[] = ['ADR1', 'EDES','PONO'];
    var title: String = 'OptiMap Script';
    var addresses: Array;
    var pbWindow,progbar,tb,tb1,wCONO,E0PA,connection,sqlQueryString,lb, ww,command, reader;
    var dbServerIp = "yourServer";
    var dbUsername = "login";
    var dbPassword = "password";
    var database = "M3FDB"+ ApplicationServices.SystemProfile.Name;
    var sqlConnectionString : String = "Driver={iSeries Access ODBC Driver};System=" + dbServerIp + ";Uid=" + dbUsername + ";Pwd=" + dbPassword + ";MGDSN=0;Initial Catalog=" + database + ";";
        
        public function Init(element : Object, args : Object, controller : Object, debug : Object) 
        {
            var content : Object = controller.RenderEngine.Content;
            this.content = content;
            this.controller = controller;
            this.debug = debug;
            wCONO = UserContext.CurrentCompany.toString();
            var n=wCONO.length;
            while(n < 3)
            {
              wCONO = "0" + wCONO ;
              n=wCONO.length;
            }
            Start();
        }
//----------------------------------------------------------------------------------------------------------------------------
public function goDLeave(sender : Object, e : DragEventArgs)
{
   if (e.Data.GetDataPresent(DataFormats.StringFormat))
   {  
      e.Effects = DragDropEffects.None;
      
      lb.Background = new SolidColorBrush(Colors.Thistle);
   }
   else
   {
      e.Effects = DragDropEffects.None;
   }
   e.Handled = true;
}
//------------------------------------------------------------------------------------------------------------------------------
public function goDEnter(sender : Object, e : DragEventArgs)
{
   //Fields for DataFormat :
   //Bitmap, CommaSeparatedValue,Dib,Dif,FileDrop,Html,Riff,Rtf,Xaml,Text,Tiff,WaveAudio...
   if (e.Data.GetDataPresent(DataFormats.StringFormat))
   {  
      e.Effects = DragDropEffects.Copy;
      
      lb.Background = new SolidColorBrush(Colors.SpringGreen);
   }
   else
   {
      e.Effects = DragDropEffects.None;
   }
   e.Handled = true;
}
//----------------------------------------------------------------------------------------------------------------------------
public function goDrop(sender : Object, e : DragEventArgs)
{
    if(e.Data.GetDataPresent(DataFormats.StringFormat))
    {
        try
        {
            var z = new Array();
            var j = 0;
            var strArr = e.Data.GetData(DataFormats.StringFormat).split('\r\n');
            for(var i =0;i<strArr.length;i++)
            {
               if(strArr[i].split(":")[0] != '' && strArr[i].split(":")[0] != 'START')
               {
                    z[j] = strArr[i].split(":")[0];
                    j++;
               }
            }
        }
        catch(ex)
        {
            debug.WriteLine(ex);
        }
        if(z.length > -1)
        {
            if(ConfirmDialog.ShowWarningDialogWithCancel("Close navigator and update those "+z.length+" deliveries ? ","Press OK to launch the job or Cancel."))
            {
                ww.Close();
                pbWindow = new Window();
                pbWindow.Title = "Progress...";
                pbWindow.Height=110;
                pbWindow.Width=300;
                pbWindow.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen;

                var myWrapPanel = new WrapPanel();
                myWrapPanel.Background = System.Windows.Media.Brushes.White;
                myWrapPanel.Orientation = Orientation.Horizontal;
                myWrapPanel.Width = 300;
                myWrapPanel.HorizontalAlignment = HorizontalAlignment.Left;
                myWrapPanel.VerticalAlignment = VerticalAlignment.Top;

                progbar = new ProgressBar();
                progbar.Width = 300;
                tb = new TextBlock();
                tb.Width = 300;
                tb.TextAlignment = "Center";
                tb.VerticalAlignment = "Center";
                tb1 = new TextBlock();

                tb1.Width = 300;
                tb1.TextAlignment = "Center";
                tb1.VerticalAlignment = "Center";

                myWrapPanel.Children.Add(progbar);
                myWrapPanel.Children.Add(tb);
                myWrapPanel.Children.Add(tb1);
                pbWindow.Content = myWrapPanel;
                pbWindow.Show();

                var worker = new BackgroundWorker();
                worker.add_DoWork(OnDoWork);
                worker.WorkerReportsProgress = true;
                worker.add_ProgressChanged(OnProgressChanged);
                worker.add_RunWorkerCompleted(OnRunWorkerCompleted); 
                worker.RunWorkerAsync(z);
            }
        }   
    }
}
//----------------------------------------------------------------------------------------------------------------------------
function Start() 
{
    try 
    {
        var listControl: MForms.ListControl = controller.RenderEngine.ListControl;
        var listView: System.Windows.Controls.ListView = controller.RenderEngine.ListViewControl;
        if (listControl == null || listView == null) 
        { 
            ConfirmDialog.ShowErrorDialog(title, 'Couldn\'t find the list.');
            return; 
        }
        
        var errors: ArrayList = new ArrayList();
        var columnIndices: ArrayList = new ArrayList();
        
        var columnDLIX: int = listControl.GetColumnIndexByName('DLIX');
        if (columnDLIX == -1) errors.Add('DLIX');
        var columnWHLO: int = listControl.GetColumnIndexByName('WHLO');
        if (columnWHLO == -1) errors.Add('WHLO');
        // find the address columns
        for (var i: int in addressFields) 
        {
            var field: String = addressFields[i];
            var columnIndex: int = listControl.GetColumnIndexByName(field);
            if (columnIndex != -1) 
            { 
                columnIndices.Add(columnIndex);
            } 
            else 
            { 
                errors.Add(field);
            }
        }
        if (errors.Count>0) 
        { 
            ConfirmDialog.ShowErrorDialog(title, 'Couldn\'t find the following columns in the list: ' + String.Join(', ', errors.ToArray()) + '.');
            return;
        }
        // get the selected rows
        var rows = listView.SelectedItems; // System.Windows.Controls.SelectedItemCollection
        if (rows == null || rows.Count == 0) 
        { 
            ConfirmDialog.ShowErrorDialog(title, 'Select one or more Deliveries in the list (press CTRL to select multiple rows).');
            return;
        }
        // get the selected addresses
        addresses = new Array();
        var k = 0;
        for (var row: ListRow in rows) 
        {
            var address: ArrayList = new ArrayList();
            for (var j: int in addressFields) 
            {
                var value: String = row[columnIndices[j]].Trim();
                if (!String.IsNullOrEmpty(value)) address.Add(value);
            }
            if (address.Count == 0) 
            { 
                ConfirmDialog.ShowErrorDialog(title, 'One or more of the selected Deliveries doesn\'t have an address.');
                return;
            }
            addresses[k] = [row[columnDLIX].Trim(),address.ToArray()];
            k=k+1;
        }
 
        // get the selected Warehouse
        var WHLO: String = rows[0][columnWHLO];
        for (row in rows) if (row[columnWHLO] != WHLO) 
        { 
            ConfirmDialog.ShowErrorDialog(title, 'The Warehouse must be the same for the selected Deliveries.');
            return;
        }
        // get the Warehouse's address
        SetWaitCursor(true);

        var record: MIRecord = new MIRecord(); 
        record['WHLO'] = WHLO;
        var parameters: MIParameters = new MIParameters(); 
        parameters.OutputFields = addressFields;
        MIWorker.Run('MMS005MI', 'GetWarehouse', record, OnCompleted, parameters);
    } 
    catch (ex: Exception) 
    {
        ConfirmDialog.ShowErrorDialog(title, ex);
        SetWaitCursor(false);
    }
}
//----------------------------------------------------------------------------------------------------------------------------
function OnCompleted(response: MIResponse) 
{
    try 
    {
        SetWaitCursor(false);
        if (response.HasError) 
        { 
            ConfirmDialog.ShowErrorDialog(title, response.ErrorMessage);
            return;
        }
        var warehouseAddress: String = '';
        for (var field in response.MetaData)
        {
            warehouseAddress += response.Item[field.Key] + ', ';
        }
        LaunchOptiMap(warehouseAddress, addresses);
    } 
    catch (ex: Exception) 
{
        ConfirmDialog.ShowErrorDialog(title, ex);
    }
}

//----------------------------------------------------------------------------------------------------------------------------
function LaunchOptiMap(fromAddress: String, toAddresses: Array) 
{
    var query: String = 'http://gebweb.net/optimap/index.php?&name0=START&loc0=' + HttpUtility.UrlEncode(fromAddress) + '&';
    // destination addresses as parameters locN + DLIX as label
    for (var i: int = 0; i < toAddresses.length; i++) 
    {
        var j = i+1;
        query += 'name'+j+'='+toAddresses[i][0]+'&loc'+j + '=' + HttpUtility.UrlEncode(toAddresses[i][1]) + '&';
    }

    // launch OptiMap in a browser
    ww = new Window();
    ww.Title = 'Optimap route optimization';
    ww.Background = new SolidColorBrush(Colors.Lavender);
    ww.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen;
    ww.WindowState = 'Maximized';

    var wb = new WebBrowser();
    
    var sp = new DockPanel();
    lb = new Label();
    lb.Content = 'Drap & drop raw path with labels here !';
    lb.HorizontalContentAlignment = 'Center';
    lb.VerticalContentAlignment = 'Center';
    lb.FontSize = 24;
    //lb.IsEnabled = false;
    lb.Background = new SolidColorBrush(Colors.Ivory);
    lb.Height = 100;
    lb.Width = 800;
    lb.AllowDrop="True";
    lb.add_DragEnter(goDEnter);
    lb.add_DragLeave(goDLeave);
    lb.add_Drop(goDrop);
    
    wb.Navigate("http://gebweb.net/optimap/index.php?"+query);
    DockPanel.SetDock(lb, Dock.Top);
    
    sp.Children.Add(lb);
    sp.Children.Add(wb);
    ww.Content = sp;
    ww.Show();
}
//----------------------------------------------------------------------------------------------------------------------------
function SetWaitCursor(wait) 
{
    
    //element.ForceCursor = true;
}
//---------------------------------------------------------------------------------------------------------------------
public function OnDoWork(sender : Object, e : DoWorkEventArgs) 
{ 
    var inputValues : Array =  e.Argument;
    var worker: BackgroundWorker = sender;
    var listDLIX = new Array();
  
    //get E0PA from CSYSTR
    worker.ReportProgress(10, "Retrieving user parameters...");
    
    connection  = new OdbcConnection(sqlConnectionString);
    sqlQueryString = "SELECT cspar1 from "+database+".CSYSTR where cscono= "+wCONO+" and csresp='"+UserContext.User+"' and cspgnm='MWS410'";
    
    try
    {
        connection.Open();
        command = new OdbcCommand(sqlQueryString, connection);
        command.CommandTimeout = 60;
        reader = command.ExecuteReader();
        if (reader.HasRows) 
        {
            while (reader.Read()) 
            {
                E0PA = reader.GetString(0).substring(49,66);
            }
            reader.Close();
        }
    }
    catch(ex)
    {
        debug.WriteLine(ex);
        reader.Close();
        connection.Close();
    }
    finally
    {
        connection.Close(); 
    }
    
    //get MSGN,IRST for DLIXs
    worker.ReportProgress(15, "Checking deliveries...");
    var wFirst = "1";
    var wList = "";
    
    
    try
    {
        connection.Open();

        for(var i=0;i<inputValues.length;i++)
        {
            sqlQueryString = "SELECT OQMSGN,OQIRST,OQDLIX from "+database+".MHDISH where oqcono= "+wCONO+" and oqinou='1' and oqdlix="+inputValues[i];
            command = new OdbcCommand(sqlQueryString, connection);
            command.CommandTimeout = 60;
            reader = command.ExecuteReader();
            if (reader.HasRows) 
            {
                while (reader.Read()) 
                {
                    listDLIX[i] = [reader.GetString(0),reader.GetString(1),reader.GetString(2)];
                }
                reader.Close();     
            }
        }
          
    }
    catch(ex)
    {
        debug.WriteLine(ex);
        reader.Close();
        connection.Close(); 
    }
    finally
    {
       connection.Close(); 
    }  
    
    
    //
    worker.ReportProgress(20, "Updating deliveries...");
    for(var i=0;i<listDLIX.length;i++)
    {
        
        if(listDLIX[i][1] == '10' || listDLIX[i][0] == '')
        {
            worker.ReportProgress(99, "Status download not OK or message number missing.");
            return;
        }
        try
        {
            var record = new MIRecord();
            record["PRFL"] = "*EXE";
            record["CONO"] = wCONO;
            record["E0PA"] = E0PA;
            record["MSGN"] = listDLIX[i][0];
            record["DLIX"] = listDLIX[i][2];
            record["INOU"] = "1";
            record["MULS"] = 1;
            record["SULS"] = i+1;

            var response = MIAccess.Execute("MYS450MI", "AddDelivery", record);

            if(!response.HasError)
            {      
              worker.ReportProgress(20+Math.round((i+1)/listDLIX.length*80),"Updating "+listDLIX[i][2]+"...");
            }
            else
            {
                debug.WriteLine(response.Error);
            }
        }
        catch(ex)
        {

        }
    }
  
}
//----------------------------------------------------------------------------------------------------------------------
 public function OnProgressChanged(sender: Object, e: ProgressChangedEventArgs) 
 {
    progbar.Value = e.ProgressPercentage;
    tb.Text =  e.ProgressPercentage + "%";    
    tb1.Text = e.UserState;
 }
 //------------------------------------------------------------------------------------------------------------------
public function OnRunWorkerCompleted(sender : Object, e : RunWorkerCompletedEventArgs) 
{
    var worker: BackgroundWorker = sender;

    if(progbar.Value != '99')
    {
        pbWindow.Close();
    }
    

    worker.remove_DoWork(OnDoWork);
    worker.remove_RunWorkerCompleted(OnRunWorkerCompleted); 
}
//---------------------------------------------------------------------------------------------------------------------

    }
}