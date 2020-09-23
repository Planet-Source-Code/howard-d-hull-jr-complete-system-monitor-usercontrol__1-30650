+=====================================+
| System Monitor ActiveX Usercontrol  |
+=====================================+

Having seen so many code snips that either only work with Windows9x or return incorrect 
values, I decided to write one that should work accross ALL version of Windows. 
I was able to find multiple pieces of code all over the place and assembled them into
a neat little ActiveX control.

This usercontrol can be be used visually to optionally display the history of any 
of the following values;
  CPU Load, Memory Load, Free Pagefile %, Free Virtual Memory % and the HD Free space %. 

While the usercontrol is Enabled the Update event is fired and additional details are passed;
  CPULoadPercent, MemoryLoadPercent, 
  PhysicalMemoryTotal, PhysicalMemoryAvailable, PhysicalMemoryAvailablePercent, 
  PageFileTotal, PageFileAvailable, PageFileAvailablePercent, 
  VirtualMemoryTotal, VirtualMemoryAvailable, VirtualMemoryAvailablePercent, 
  HDTotalBytes, HDTotalFreeBytes, HDAvailableFreeBytes, HDTotalBytesUsed, HDAvailablePercent

The control can also be hidden and used to manually retreive the above values using 
the GetCurrentSystemLevels routine. This allows you to retreive current system levels and 
display the information using your own layout/design.

The visual look of the control is fully customizable. All colors can be modified to fit the
individual needs of developers. 




Note:
The difference between HDAvailableFreeBytes and HDTotalFreeBytes;

HDTotalFreeBytes
  The total number of free bytes on the disk.

HDAvailableFreeBytes
  The total number of free bytes on the disk that are available 
  to the user associated with the calling thread.
  If per-user quotas are in use, this value may be less than 
  the HDTotalFreeBytes value.


Contributing Code Authors:
The KPD-Team at AllApi (http://www.allapi.net)
Randy Birch  at VBnet  (http://www.mvps.org/vbnet/index.html)
Sami Riihilahti at PSC (http://www.planet-source-code.com/xq/ASP/txtCodeId.14620/lngWId.1/qx/vb/scripts/ShowCode.htm)
Various Usenet posters
Anonymous Code posters
