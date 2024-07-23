rule KNOCKOUT
{
  meta:
    description = "Detects the pre-compiled artifact collection tool 'KNOCKOUTx64.exe'"
    author = "Steffen Rogge / NVISO"
    date = "2024-07-18"
    reference = "https://blog.nviso.eu/"
    hash = "3B248C3E8B64719D5991A762330A0B2BB58E116247DA89E781EC1A53F4ED1D00"
    hash = "86C0199B6A9305621B011DAAF999A4FE0F266BA9"
    hash = "58B142287E47B5605363639F5C4ABB45"
  strings:
    $guid = "ad5cead9-cde0-4362-83eb-7d80de0025c9"
    $toolname = "KNOCKOUT"
    $pdb = "KNOCKOUT.pdb"
    $applistresource = "KNOCKOUT.AppIdlist.csv"
    $function1 = "ExtractRecentLNKFilesAndFolders"
    $function2 = "ExtractRecentURLFilesAndFolders"
    $function3 = "ExtractRecentOfficeFilesAndFolders"
    $function4 = "ExtractRecentExplorerFilesAndFolders"
    $function5 = "ExtractFrequentFilesFromJumpLists"
  condition:
    3 of them
}
