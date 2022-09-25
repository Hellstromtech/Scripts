$revPos = '670BA-XX-XXX'.Length
dir '670BA*.pdf' | group @{e={ $_.Name.Substring(0, $revPos) }} | % {
    $revs = $_.Group | % { $_.Name.Substring($revPos).Split('.')[0] } | group Length | sort -Descending -Property @{e={ [int]$_.Name }} | % { $_.Group | sort -Descending }
    $fileSet = $_.Name
    $revs | % { $fileSet + $_ + '.pdf' } | select -Skip 1 | del
}