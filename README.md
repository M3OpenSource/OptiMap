# OptiMap
Route optimization for MWS410 with OptiMap

* [OptiMap_V1](https://m3ideas.org/2012/10/04/route-optimization-for-mws410-with-optimap/)
* [OptiMap_V2](https://m3ideas.org/2013/03/08/optimap_v2/)

## V4 [PENDING]
* Added ability to Export OptiMap's route to M3 Loading (MULS) and Unloading sequence (SULS) using API MYS450MI.AddDelivery, closing the loop of integrating M3 to OptiMap.

## V3
* Made Warehouse dynamic: the Warehouse (WHLO) is now detected from the selected rows, and its address is dynamically retrieved using API MMS005MI.GetWarehouse; no need to specify the Warehouse address as a parameter anymore.
* Improved error messages
* Made addressFields optional as parameter of the script
* Added Delegate

## V2
* Can now specify an optional Warehouse address as Argument of the script as the starting and ending points of the delivery route.

## V1
* First version of the script, gets the ADR1, ADR2, and ADR3 of the selected rows, and sends to http://www.optimap.net/
