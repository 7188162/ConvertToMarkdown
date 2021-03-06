<#
.SYNOPSIS
オブジェクトをMarkdownに変換する

.DESCRIPTION
オブジェクトをMarkdownに変換します。列順は順不同です。

.PARAMETER Object
Object - パイプライン入力可

.EXAMPLE
$arrayObjects | ConvertTo-Markdown

.NOTES

#>
function ConvertTo-Markdown {
    Param(
    [parameter(mandatory=$true,
    ValueFromPipeline=$true)] $Object
    )
    
    begin {        
        $lines=@()   
    }
    
    process {               
        $item=@()    
        # ヘッダの取得
        $props = $Object | Get-Member -MemberType Properties | Select-Object -ExpandProperty Name
        # 中身の取得                
        $props | % {        
            $prop = $_       
            $item += $Object.$prop        
            $line = $item -join '|'                    
        }   
        $lines += ("|" + $line + "|")     
    }
    end {       
        # ヘッダの出力        
        $str="|"+($props -join '|')+"|"
        Write-Output $str
        
        # ヘッダ区切りの出力
        $strSeparator=""
        for ($i=1;$i -le $props.count;$i++) {
            $strSeparator += ("|---")
        }
        
        $strSeparator += "|"
        Write-Output $strSeparator
        $lines | ForEach-Object {Write-Output $_}
    }
}