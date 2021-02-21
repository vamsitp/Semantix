# simMatch
dotnet tool to find similarities b/w 2 sets of _Titles_

![Screenshot](https://github.com/vamsitp/SimMatch/blob/master/Screenshot.png?raw=true)

output:
![Output_Screenshot](https://github.com/vamsitp/SimMatch/blob/master/Output_Screenshot.png?raw=true)

**installation** (_pre-req_: [`dotnet 5`](https://dotnet.microsoft.com/download/dotnet/5.0))
> `dotnet tool install -g --ignore-failed-sources SimMatch`   

**usage**
> `match` [space] `"path-to-excel-file"` [space] `min similarity threshold (b/w 0.0 - 1.0)` [space] `max top similarities`   
> **e.g.** `match` `"D:\List1_List2.xlsx"` `0.4` `1`   
> **note**: Input Excel must contain 2 sheets with 2 columns (named `ID` & `Title`) or at least 1 column (named `Title`) 

**credits**
- [Tyler Jensen - Keyword extraction in C# with Word-co-occurrence algorithm](https://www.tsjensen.com/blog/post/2010/03/14/Keyword+Extraction+In+C+With+Word+Cooccurrence+Algorithm)
- [SimMetrics.NET](https://github.com/StefH/SimMetrics.Net/#simmetricsnet)