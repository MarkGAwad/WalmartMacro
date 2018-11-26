# WalmartMacro
Walmart Import and export Macro
# V 1.0.1
## Updated 11/26/2018
Added the feature of calculating the volume of every PO

This is based on the following information

Item 	               Cube

BYT2003SPV	         1.2413

BT301SSL	           1.4583

CL-CB-SV-100-GY    	1.515

BY-ST-RG-100-BK    	0.0729

BY-PP-XL-101-BK   	1.4068

BY-PP-LG-102-BK	    1.4583

BY-SP-UN-100-CR	    0.037

CL-PG-KT-101-BK   	1.6042

BY-PP-XL-102-BK   	1.515

BY-PP-LG-114-AC	    1.145

The equation is to search by item ID ( get the Array third dimension ) and multiply it with the number of Cartons.

Add it to the order volume in the loop and then print it in Excel Cell"AD".

Also do the same with Back orders in a loop.

