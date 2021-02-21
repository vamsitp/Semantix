# simMatch
dotnet tool to find similarities/matches between 2 sets of _Titles_ / _Phrases_ using the chosen _algorithms_ (with _StopWords_ and _Stemming_)

![Screenshot](https://github.com/vamsitp/SimMatch/blob/master/Screenshot.png?raw=true)

output:
![Output_Screenshot](https://github.com/vamsitp/SimMatch/blob/master/Output_Screenshot.png?raw=true)

**installation** (_pre-req_: [`dotnet 5`](https://dotnet.microsoft.com/download/dotnet/5.0))
> `dotnet tool install -g --ignore-failed-sources SimMatch`   

**usage**
> `match` [space] `"path-to-excel-file"` [space] `min similarity threshold (b/w 0.0 - 1.0)` [space] `max top similarities`   
> **e.g.** `match` `"D:\List1_List2.xlsx"` `0.25` `2`   
> **note**: Input Excel must contain 2 sheets with 2 columns (named `ID` & `Title`) or at least 1 column (named `Title`) 

**credits**
- [SimMetrics.NET](https://github.com/StefH/SimMetrics.Net/#simmetricsnet)
- [Annytab.Stemmer](https://github.com/annytab/a-stemmer)
- [dotnet-stop-words](https://github.com/hklemp/dotnet-stop-words)