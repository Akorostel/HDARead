# HDARead
Command line OPC HDA client (.NET)

HDARead is used to read the data from OPC HDA server and save it to text file.

## Examples

HDARead.exe -s="OPC.PHDServerHDA.1" $TEST.TEST2.QQ.LOLIM
HDARead.exe -s=OPC.PHDServerHDA.1 -from="09/17/19 11:00 AM" -to=NOW -a=START -r=600 $TEST.TEST2.QQ.LOLIM
HDARead.exe -s=OPC.PHDServerHDA.1 -from="09/17/19 11:00 AM" -to=NOW -a=INTERPOLATIVE -r=600 $TEST.TEST2.QQ.LOLIM $TEST.TEST2.QQ.INPUT
HDARead.exe -s=OPC.PHDServerHDA.1 -from="09/17/19 15:00" -to=NOW -a=INTERPOLATIVE -r=600 $TEST.TEST2.QQ.LOLIM $TEST.TEST2.QQ.HILIM > out.txt
HDARead.exe -s=OPC.PHDServerHDA.1 -from="09/17/19 15:00" -to=NOW -a=INTERPOLATIVE -r=600 -f="yyyy-MM-dd HH-mm-ss" $TEST.TEST2.QQ.LOLIM $TEST.TEST2.QQ.HILIM > out.txt
HDARead.exe -s=OPC.PHDServerHDA.1 -from="09/17/19 15:00" -to=NOW -a=INTERPOLATIVE -r=600 -f="yyyy-MM-dd HH-mm-ss" $TEST.TEST2.QQ.LOLIM $TEST.TEST2.QQ.HILIM -o="out.txt"
HDARead.exe -s=OPC.PHDServerHDA.1 -from="09/19/19 15:00" -to=NOW -a=INTERPOLATIVE -r=600 -f=TABLE $TEST.TEST2.QQ.LOLIM $TEST.TEST2.QQ.HILIM -o="out.csv"
HDARead.exe -s=OPC.PHDServerHDA.1 -from="09/19/19 15:00" -to=NOW -a=INTERPOLATIVE -r=600 -f=TABLE -i=tags.txt -o="out.csv"
HDARead.exe -s=OPC.PHDServerHDA.1 -from="09/24/19 10:00" -to="09/24/19 18:00" -raw  -f=TABLE $TEST.TEST2.QQ.LOLIM $TEST.TEST2.QQ.HILIM TEST.DATA.MI2

## Start and end time format

## Modes: ReadRaw and ReadProcessed

## Aggregates

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

## Output formats: LIST, TABLE, MERGED

## Arguments

Usage: HDARead [OPTIONS]+ tag1 tag2 tag3 ...

Options:
  -n=VALUE, --node=VALUE
        
        Remote computer name (optional)
        
  -s=VALUE, --server=VALUE

        OPC HDA server name (required)

  --from=VALUE, --start=VALUE, --begin=VALUE

        Start time (abs. or relative), default NOW-1H

  --to=VALUE, --end=VALUE
    
        End time (abs. or relative), default NOW

  -a=VALUE, --agg=VALUE
  
        Aggregate for ReadProcessed (see spec)
   
   -r=VALUE, --resample=VALUE    
        
        Resample interval for ReadProcessed (in seconds), 0 - return just one value (see OPC HDA spec.)
        
   --raw
   
        Read raw data (if omitted, read processed data)
   
   -m=VALUE, maxvalues=VALUE
   
        Maximum number of values to load for each tag (only for ReadRaw data)

  -t=VALUE, --tsformat=VALUE
  
        Output timestamp format to use
  
  -f=VALUE
  
        Output format: LIST (default) or TABLE or MERGED
        
  -q=VALUE
  
        Include quality in output data: NONE (default), DA, HISTORIAN or BOTH
        
  -o=VALUE, --output=VALUE

        Output filename (if omitted, output to console)
        
  -i=VALUE, --input=VALUE  
  
        Input filename with list of tags (if omitted, tag list must be provided as command line argument)
        
   -v, --verbose
   
        Output debug info
        
  -h, -?, --help

        Show usage info


