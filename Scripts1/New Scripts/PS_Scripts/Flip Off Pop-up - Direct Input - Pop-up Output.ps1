$Message = "
                /'''\
               |\__/|
               |       |
               |       |
               |>-<|
               |       |
        /'''\|        |/'''\...
 /''\|       |       |       |   \
|      |       |       |       |    \
|      |       |       |       |      \
|  ~     ~     ~     ~  |)       )
|                                      /
 \                                   /
   \                               /
     \                           /
      |                         |
      |                         |
"
$c = New-Object -Comobject wscript.shell
$b = $c.popup("$Message",0,"325 CS Wins!",0)