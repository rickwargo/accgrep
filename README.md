accgrep
======

**Grep** for **Access** files

This command line utility will perform a *regular expression* search on many
Access objects. It is only meant to be run from *Windows* as it manipulates the
Access COM object through **Win32ole automation**. 

It decomposes all of the significant objects in Excel into text streams and
applies regex matching to find items of interest. It is also capable of 
find and replace.

```
Usage: accgrep [options] [expression] file ...

Options:
    -l, --files-with-matches         only print file names containing matches
    -L, --files-without-matches      only print file names containing no match
    -i, --ignore-case                ignore case distinctions
    -X, --extended                   use extended regular expressions
    -M, --multi-line                 search across lines
    -v, --invert-match               select non-matching lines
    -n, --line-numbers               print line number with output lines
        --recurse                    recurse into directories
    -e, --regexp PATTERN             use PATTERN as a regular expression (may have multiples)
    -r, --replace STRING             use STRING as a replacement to the regular expression
    -D, --delete-matching-line       delete lines matching the regular expression
    -s, --search WHAT                database objects to search
                                       (all, macros, procedures, references, forms, reports, tables, queries, data)
    -c, --controls NAME              search only controls matching NAME (should include property)
    -p, --properties NAME            search only properties matching NAME (should include form or report name)
    -f, --field NAME                 search only field named NAME
    -P, --procedure NAME             search only the NAMEd procedure (may have multiples)
    -m, --max-count NUM              stop after NUM matches
        --include PATTERN            files that match PATTERN will be examined
        --exclude PATTERN            files that match PATTERN will be skipped
    -F, --forms-matching PATTERN     forms that match PATTERN will be examined
    -Q, --queries-matching PATTERN   queries that match PATTERN will be examined
    -R, --reports-matching PATTERN   reports that match PATTERN will be examined
    -T, --tables-matching PATTERN    tables that match PATTERN will be examined
        --linked-tables              only search linked tables (Connect string is not empty)
    -w, --where CLAUSE               use clause to limit rows in searching table data
    -a, --and                        specify AND between the where clauses
    -o, --or                         specify OR between the where clauses

Options that shouldn't be options:
    -C, --recycle-every NUM          recycle access application every NUM times

Common options:
    -h, --help                       show this message
    -V, --verbose                    show messages indicating progress
        --version                    show version
```
