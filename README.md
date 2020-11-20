HDARead
=======

Command line OPC HDA client (.NET)

HDARead is used to read the data from OPC HDA server and save it to text file.

Examples
--------

~~~~
HDARead.exe -s=OPCServerHDA.1 Tag1

HDARead.exe -s=OPCServerHDA.1 -from="09/17/19 11:00 AM" -to=NOW -a=START -r=600 Tag1

HDARead.exe -s=OPCServerHDA.1 -from="09/17/19 11:00 AM" -to=NOW -a=INTERPOLATIVE -r=600 Tag1 Tag2

HDARead.exe -s=OPCServerHDA.1 -from="09/17/19 15:00" -to=NOW -a=INTERPOLATIVE -r=600 Tag1 Tag2 > out.txt

HDARead.exe -s=OPCServerHDA.1 -from="09/17/19 15:00" -to=NOW -a=INTERPOLATIVE -r=600 -f="yyyy-MM-dd HH-mm-ss" Tag1 Tag2 > out.txt

HDARead.exe -s=OPCServerHDA.1 -from="09/17/19 15:00" -to=NOW -a=INTERPOLATIVE -r=600 -f="yyyy-MM-dd HH-mm-ss" Tag1 Tag2 -o="out.txt"

HDARead.exe -s=OPCServerHDA.1 -from="09/19/19 15:00" -to=NOW -a=INTERPOLATIVE -r=600 -f=TABLE Tag1 Tag2 -o="out.csv"

HDARead.exe -s=OPCServerHDA.1 -from="09/19/19 15:00" -to=NOW -a=INTERPOLATIVE -r=600 -f=TABLE -i=tags.txt -o="out.csv"

HDARead.exe -s=OPCServerHDA.1 -from="09/24/19 10:00" -to="09/24/19 18:00" -raw  -f=TABLE Tag1 Tag2 Tag3
~~~~

Start and end time format
-------------------------

Start and end timestamps may be specified as either absolute or relative. Absolute timestamps are parsed by .NET System.DateTime.Parse function. See [description](https://docs.microsoft.com/en-us/dotnet/api/system.datetime.parse). Relative timestamps are in the form `keyword+/-offset+/-offset…`, i.e. `NOW-1D+7H30M`. For details see OPCHDA_TIME description in OPC HDA specs.

Modes: ReadRaw and ReadProcessed
--------------------------------

Data can be queried from server by either ReadRaw or ReadProcessed functions. Basically, ReadRaw reads data as they are stored on server, while ReadProcessed divide the `[StartTime, EndTime)` range to intervals of the length `ResampleInterval` and applies selected aggregate function to every time interval. This processing is done by OPC HDA server. For details see OPC HDA specs.

Aggregates
----------

- ANNOTATIONS = 24,
- WORSTQUALITY = 23,
- PERCENTBAD = 22,
- PERCENTGOOD = 21,
- DURATIONBAD = 20,
- DURATIONGOOD = 19,
- RANGE = 18,
- VARIANCE = 17,
- REGDEV = 16,
- REGCONST = 15,
- REGSLOPE = 14,
- DELTA = 13,
- END = 12,
- START = 11,
- MAXIMUM = 10,
- MAXIMUMACTUALTIME = 9,
- MINIMUM = 8,
- MINIMUMACTUALTIME = 7,
- STDEV = 6,
- COUNT = 5,
- TIMEAVERAGE = 4,
- AVERAGE = 3,
- TOTAL = 2,
- INTERPOLATIVE = 1,
- NOAGGREGATE = 0

Output formats: TABLE, MERGED
-----------------------------

HDAread can show queried data on console (if no `-o` key specified) or save them to text file (CSV - comma separated value). 

There are two output formats supported: TABLE and MERGED.  
TABLE formatting looks like the following:  
`Tag1Timestamp, Tag1Value, Tag2Timestamp, Tag2Value, Tag3Timestamp, Tag3Value...`

MERGE formatting looks like the following:  
`Timestamp, Tag1Value, Tag2Value, Tag3Value...`

When querying *raw* data from server, server returns individual set of timestamps for each tag, so additional processing is done by HDAread to align different tags to single timeline for MERGED output. 

There is also an option `-q` to include in the output the quality for each tag.

Output timestamp format ("yyyy-MM-dd hh:mm:ss", "MM/dd/yy hh:mm", etc.) can be specified by using `-t` key. For details on format strings see [1](https://docs.microsoft.com/en-us/dotnet/standard/base-types/standard-date-and-time-format-strings), [2](https://docs.microsoft.com/en-us/dotnet/standard/base-types/custom-date-and-time-format-strings)
One special format string is -t=DateTime, which outputs date and time in separate columns. The format is fixed: MM/dd/yyyy,HH:mm:ss
 
Arguments
---------

Usage: HDARead [OPTIONS]+ tag1 tag2 tag3 ...

Options:  
  -n=VALUE, --node=VALUE
    
        Remote computer name (optional)
        
  -s=VALUE, --server=VALUE

        OPC HDA server name (required)

  --from=VALUE, --start=VALUE, --begin=VALUE

        Start time (abs. or relative), default is NOW-1H

  --to=VALUE, --end=VALUE
    
        End time (abs. or relative), default is NOW

  -a=VALUE, --agg=VALUE
  
        Aggregate for ReadProcessed (see spec)
   
   -r=VALUE, --resample=VALUE    
        
        Resample interval for ReadProcessed (in seconds), 0 - return just one value (see OPC HDA spec.)
        
   --raw
   
        Read raw data (if omitted, read processed data)
   
   -m=VALUE, maxvalues=VALUE
   
        Maximum number of values to load for each tag (only for ReadRaw)
   
   -b, --bounds
        
        Whether the bounding item values should be returned (only for ReadRaw)

  -t=VALUE, --tsformat=VALUE
  
        Output timestamp format to use
  
  -f=VALUE
  
        Output format: TABLE or MERGED (default)
        
  -q=VALUE
  
        Include quality in output data: NONE (default), DA, HISTORIAN or BOTH
        
  -o=VALUE, --output=VALUE

        Output filename (if omitted, output to console)
        
  -i=VALUE, --input=VALUE  
  
        Input filename with list of tags (if omitted, tag list must be provided as command line argument)
        
  -v
   
        Show extended info


  -vv, --verbose
   
        Show debug info
        
  -h, -?, --help

        Show usage info


Third-party libraries used in this product
-----------------------------------------
This product uses the following software:
 - [NDesk Options:  Copyright (C) 2008 Novell](http://www.ndesk.org/Options)
 - [OPC Redistributables](https://opcfoundation.org/developer-tools/samples-and-tools-classic/core-components/)