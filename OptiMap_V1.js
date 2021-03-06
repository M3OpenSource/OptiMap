import System;
import System.Web;
import System.Windows;
import Mango.UI.Services.Lists;
import MForms;
 
/*
Integrates the Smart Office Delivery Toolbox - MWS410/B with OptiMap - Fastest Roundtrip Solver, http://www.optimap.net/
to calculate and show on Google Maps the fastest roundtrip for the selected Routes.
This is interesting to reduce driving time and cost, and for a driver to optimize its truck load according to the order of delivery.
1) Deploy this script in the mne\jscript\ folder in your Smart Office server
2) Create a Shortcut in MWS410/B to run this script; for that go to MWS410/B > Tools > Personalize > Shortcuts > Advanced > Script Shortcut, set the Name to OptiMap, and set the Script name to OptiMap
3) Create a View (PAVR) in MWS410/B that shows the address columns ADR1, ADR2, ADR3
4) Select multiple Routes in the list (press CTRL to select multiple rows)
5) Click the OptiMap Shortcut to run this script and launch OptiMap for the selected Routes
For more information and screenshots refer to https://thibaudatwork.wordpress.com/2012/10/04/route-optimizer/
Thibaud Lopez Schneider, October 4, 2012 (rev.2)
*/
package MForms.JScript {
    class OptiMap_V1 {
        public function Init(element: Object, args: Object, controller : Object, debug : Object) {
            try {
                // get the list
                var listControl: MForms.ListControl = controller.RenderEngine.ListControl;
                var listView: System.Windows.Controls.ListView = controller.RenderEngine.ListViewControl;
                if (listControl == null || listView == null) { MessageBox.Show('Error: Couldn\'t find the list.'); return; }
                // get the selected rows
                var rows = listView.SelectedItems; // System.Windows.Controls.SelectedItemCollection
                if (rows == null || rows.Count == 0) { MessageBox.Show('Error: No rows selected.'); return; }
                // get the address columns ADR1, ADR2, ADR3
                var column1: int = listControl.GetColumnIndexByName('ADR1');
                var column2: int = listControl.GetColumnIndexByName('ADR2');
                var column3: int = listControl.GetColumnIndexByName('ADR3');
                if (column1 == -1 || column2 == -1 || column3 == -1) { MessageBox.Show('Error: Couldn\'t find the address columns ADR1, ADR2, ADR3.'); return; }
                // construct the URL
                var query: String = '';
                for (var i: int = 0; i < rows.Count; i++) {
                    var row: ListRow = rows[i];
                    var ADR1: String = row[column1];
                    var ADR2: String = row[column2];
                    var ADR3: String = row[column3];
                    var address: String = ADR1 + ',' + ADR2 + ',' + ADR3;
                    query += 'loc' + i + '=' + HttpUtility.UrlEncode(address) + '&';
                }
                var uri: Uri = new Uri('http://www.optimap.net/?' + query);
                // launch OptiMap
                ScriptUtil.Launch(uri);
            } catch (ex: Exception) {
                MessageBox.Show(ex);
            }
        }
    }
}
