$wrd = new-object -com word.application
$wrd.visible = $false

$opt = $wrd.options

$opt.replaceSelection                             = $true    #   Typing replaces selected text
$opt.autoWordSelection                            = $true    #   When selecting, automatically select entire word
$opt.allowDragAndDrop                             = $true    #   Allow text to be dragged and dropped
$opt.ctrlClickHyperlinkToOpen                     = $true    #   Use CTRL + Click to follow hyperlink
$opt.smartParaSelection                           = $true    #   Use smart paragraph selection
$opt.smartCursoring                               = $true    #   Use smart cursoring
$opt.INSKeyForOvertype                            = $true    #   Use the Insert key to control overtype mode (overtype must be true to have effect).
$opt.overtype                                     = $true    #       Use overtype mode
$opt.promptUpdateStyle                            = $true    #   Prompt to update style
$opt.useNormalStyleForList                        = $true    #   Use normal style for bulleted or numbered list
$opt.formatScanning                               = $false   #   Keep track of formatting
$opt.showFormatError                              = $false   #       Mark formatting inconsistencies
$opt.updateStyleListBehavior                      =  0       #   Updating style to match selection                ( 0 = Keep previous numbering and bullets pattern ; 1 = Add numbering or bullets to all paragraphs with this style)
$opt.allowClickAndTypeMouse                       = $true    #   Enable click and type
#  ????                                                      #       Default paragraph style
#  ????                                                      #   Show AutoComplete suggestions
#  ????                                           = $true    #   Do not automatically hyperlink screenshot
$opt.autoKeyboardSwitching                        = $true    #   Automatically switch keyboard to match language of surrounding text

$wrd.quit()
