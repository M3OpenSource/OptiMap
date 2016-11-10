import System;
import System.Collections;
import System.Windows.Input;
import System.Windows.Threading;
import System.Web;
import System.Windows;
import Lawson.M3.MI;
import Mango.UI;
import Mango.UI.Services.Lists;
import MForms;

/*

    OptiMap for M3
    Thibaud Lopez Schneider, October 19, 2012 (rev.3)

    This script illustrates how to integrate the Smart Office Delivery Toolbox - MWS410/B with OptiMap - Fastest Roundtrip Solver, http://www.optimap.net/
    to calculate and show on Google Maps the fastest roundtrip from the Warehouse to the selected Delivery addresses; it's an application of the Traveling Salesman Problem (TSP) to M3.
    This is interesting for a company to reduce overall driving time and cost, and for a driver to optimize its truck load according to the order of delivery.

    To install this script:
    1) Deploy this script in the mne\jscript\ folder of your Smart Office server
    2) Create a Shortcut in MWS410/B to run this script; for that go to MWS410/B > Tools > Personalize > Shortcuts > Advanced > Script Shortcut, set the Name to OptiMap_V3, and set the Script name to OptiMap_V3
    3) Optionnally, set the address field names as the Script argument, for example: "ADR1, ADR2, ADR3, ADR4", or "TOWN, ECAR, PONO, CSCD"; by default it's "ADR1, ADR2, ADR3"
    3) Use or create a View (PAVR) in MWS410/B that shows the Warehouse WHLO and the specified address fields

    To use this script:
    1) Start MWS410/B
    2) Switch to the View (PAVR) that shows the Warehouse WHLO and the specified address fields
    3) Select one or more Deliveries in the list (press CTRL to select multiple rows)
    4) Click the OptiMap Shortcut
    5) The Shortcut will run the script, the script will get the Warehouse address by calling MMS005MI.GetWarehouse, will launch OptiMap in a browser and pass the selected Delivery addresses as locN parameters in the URL, and OptiMap will optimize the roundtrip

    For more information and screenshots refer to http://thibaudatwork.wordpress.com/2012/10/04/route-optimization-for-mws410-with-optimap/

*/
package MForms.JScript {
    class OptiMap_V3 {

        /*
            Default settings, change to suit your needs
        */
        var addressFields: String[] = ['ADR1', 'ADR2', 'ADR3'];

        /*
            Global variables
        */
        var title: String = 'OptiMap Script';
        var controller: MForms.InstanceController;
        var addresses: ArrayList;

        /*
            Main entry point
        */
        public function Init(element: Object, args: Object, controller : Object, debug : Object) {
            try {
                this.controller = controller;
                // optionnally, get the list of address fields from the argument
                if (!String.IsNullOrEmpty(args)) {
                    addressFields = args.Replace(' ', '').Split(',');
                }
                // wait for UI to render, then start
                var StartDelegate: Action = Start;
                controller.RenderEngine.Content.Dispatcher.BeginInvoke(DispatcherPriority.Background, StartDelegate);
            } catch (ex: Exception) {
                ConfirmDialog.ShowErrorDialog(title, ex);
            }
        }

        /*
            Start
        */
        function Start() {
            try {
                // get a reference to the list
                var listControl: MForms.ListControl = controller.RenderEngine.ListControl;
                var listView: System.Windows.Controls.ListView = controller.RenderEngine.ListViewControl;
                if (listControl == null || listView == null) { ConfirmDialog.ShowErrorDialog(title, 'Couldn\'t find the list.'); return; }
                // get the columns indices
                var errors: ArrayList = new ArrayList();
                var columnIndices: ArrayList = new ArrayList();
                // find the Warehouse column
                var columnWHLO: int = listControl.GetColumnIndexByName('WHLO');
                if (columnWHLO == -1) errors.Add('WHLO');
                // find the address columns
                for (var i: int in addressFields) {
                    var field: String = addressFields[i];
                    var columnIndex: int = listControl.GetColumnIndexByName(field);
                    if (columnIndex != -1) { columnIndices.Add(columnIndex); } else { errors.Add(field); }
                }
                if (errors.Count>0) { ConfirmDialog.ShowErrorDialog(title, 'Couldn\'t find the following columns in the list: ' + String.Join(', ', errors.ToArray()) + '.'); return; }
                // get the selected rows
                var rows = listView.SelectedItems; // System.Windows.Controls.SelectedItemCollection
                if (rows == null || rows.Count == 0) { ConfirmDialog.ShowErrorDialog(title, 'Select one or more Deliveries in the list (press CTRL to select multiple rows).'); return; }
                // get the selected addresses
                addresses = new ArrayList();
                for (var row: ListRow in rows) {
                    var address: ArrayList = new ArrayList();
                    for (var j: int in addressFields) {
                        var value: String = row[columnIndices[j]].Trim();
                        if (!String.IsNullOrEmpty(value)) address.Add(value);
                    }
                    if (address.Count == 0) { ConfirmDialog.ShowErrorDialog(title, 'One or more of the selected Deliveries doesn\'t have an address.'); return; }
                    addresses.Add(String.Join(', ', address.ToArray()));
                }
                // get the selected Warehouse
                var WHLO: String = rows[0][columnWHLO];
                for (row in rows) if (row[columnWHLO] != WHLO) { ConfirmDialog.ShowErrorDialog(title, 'The Warehouse must be the same for the selected Deliveries.'); return; }
                // get the Warehouse's address
                SetWaitCursor(true);
                var record: MIRecord = new MIRecord(); record['WHLO'] = WHLO;
                var parameters: MIParameters = new MIParameters(); parameters.OutputFields = addressFields;
                MIWorker.Run('MMS005MI', 'GetWarehouse', record, OnCompleted, parameters);
            } catch (ex: Exception) {
                ConfirmDialog.ShowErrorDialog(title, ex);
                SetWaitCursor(false);
            }
        }
        function OnCompleted(response: MIResponse) {
            try {
                SetWaitCursor(false);
                if (response.HasError) { ConfirmDialog.ShowErrorDialog(title, response.ErrorMessage); return; }
                var warehouseAddress: String = '';
                for (var field in response.MetaData) warehouseAddress += response.Item[field.Key] + ', ';
                LaunchOptiMap(warehouseAddress, addresses);
            } catch (ex: Exception) {
                ConfirmDialog.ShowErrorDialog(title, ex);
            }
        }

        /*
            Launches OptiMap in a browser for the specified starting address, to the specified destination addresses.
        */
        function LaunchOptiMap(fromAddress: String, toAddresses: ArrayList) {
            // URL with starting address loc0
            var query: String = 'http://www.optimap.net/?loc0=' + HttpUtility.UrlEncode(fromAddress) + '&';
            // destination addresses as parameters locN
            for (var i: int = 0; i < toAddresses.Count; i++) {
                query += 'loc'+(i+1) + '=' + HttpUtility.UrlEncode(toAddresses[i]) + '&';
            }
            // launch OptiMap in a browser
            ScriptUtil.Launch(new Uri(query));
        }

        /*
            Shows|hides the wait cursor.
        */
        function SetWaitCursor(wait) {
            var element = controller.Host.Implementation;
            element.Cursor = wait ? Cursors.Wait : Cursors.Arrow;
            //element.ForceCursor = true;
        }

    }
}
