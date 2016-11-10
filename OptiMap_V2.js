import System;
import System.Web;
import System.Windows;
import Mango.UI.Services.Lists;
import MForms;

/*

    OptiMap_V2 for M3
    Thibaud Lopez Schneider, October 19, 2012 (rev.2)

    This script illustrates how to integrate the Smart Office Delivery Toolbox - MWS410/B with OptiMap - Fastest Roundtrip Solver, http://www.optimap.net/
    to calculate and show on Google Maps the fastest roundtrip for the selected Routes; it's an application of the Traveling Salesman Problem (TSP) to M3.
    This is interesting for a company to reduce overall driving time and cost, and for a driver to optimize its truck load according to the order of delivery.
 
    To install this script:
    1) Deploy this script in the mne\jscript\ folder of your Smart Office server
    2) Create a Shortcut in MWS410/B to run this script; for that go to MWS410/B > Tools > Personalize > Shortcuts > Advanced > Script Shortcut, set the Name to OptiMap, and set the Script name to OptiMap
    3) Optionally, set the starting address (for example the Warehouse) as the Script argument; the address must be recognized by Google Maps
    4) Create a View (PAVR) in MWS410/B that shows the address columns ADR1, ADR2, ADR3
 
    To use this script:
    1) Start MWS410/B
    2) Switch to the View (PAVR) that shows the address columns ADR1, ADR2, ADR3
    3) Select multiple Routes in the list (press CTRL to select multiple rows)
    4) Click the OptiMap Shortcut
    5) The Shortcut will run the script, the script will launch OptiMap in a browser and pass the selected addresses as locN parameters in the URL, and OptiMap will optimize the roundtrip
 
    For more information and screenshots refer to:
    https://thibaudatwork.wordpress.com/2012/10/04/route-optimization-for-mws410-with-optimap/
    https://thibaudatwork.wordpress.com/2013/03/08/optimap_v2/

*/
package MForms.JScript {
    class OptiMap_V2 {
        public function Init(element: Object, args: Object, controller : Object, debug : Object) {
            try {
                // get a reference to the list
                var listControl: MForms.ListControl = controller.RenderEngine.ListControl;
                var listView: System.Windows.Controls.ListView = controller.RenderEngine.ListViewControl;
                if (listControl == null || listView == null) { MessageBox.Show('Error: Couldn\'t find the list.'); return; }
                // get the selected rows
                var rows = listView.SelectedItems; // System.Windows.Controls.SelectedItemCollection
                if (rows == null || rows.Count == 0) { MessageBox.Show('Error: Select multiple routes in the list (press CTRL to select multiple rows).'); return; }
                // get the address columns ADR1, ADR2, ADR3
                var column1: int = listControl.GetColumnIndexByName('ADR1');
                var column2: int = listControl.GetColumnIndexByName('ADR2');
                var column3: int = listControl.GetColumnIndexByName('ADR3');
                if (column1 == -1 || column2 == -1 || column3 == -1) { MessageBox.Show('Error: Couldn\'t find the address columns ADR1, ADR2, ADR3.'); return; }
                // construct the URL
                var query: String = 'http://www.optimap.net/?';
                // set the optional starting address
                var offset: int = 0;
                if (!String.IsNullOrEmpty(args)) { offset=1; query += 'loc0=' + HttpUtility.UrlEncode(args) + '&'; }
                // add the selected addresses
                for (var i: int = 0; i < rows.Count; i++) {
                    var row: ListRow = rows[i];
                    var ADR1: String = row[column1];
                    var ADR2: String = row[column2];
                    var ADR3: String = row[column3];
                    var address: String = ADR1 + ',' + ADR2 + ',' + ADR3;
                    query += 'loc' + (i+offset) + '=' + HttpUtility.UrlEncode(address) + '&';
                }
                // launch OptiMap in a browser
                ScriptUtil.Launch(new Uri(query));
            } catch (ex: Exception) {
                MessageBox.Show(ex);
            }
        }
    }
}
