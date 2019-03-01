#
#
#                                         Microsoft Word Proofing options:
#
#                                         Change how Word corrects and formats your text.
#
#

$wrd = new-object -com word.application
$wrd.visible = $false

$opt = $wrd.options
$auc = $wrd.autoCorrect


                                                             #   A u t o C o r r e c t   O p t i o n s
                                                             #   =====================================
                                                             #
                                                             #   AutoCorrect
                                                             #   --------------------------------------------------
$auc.displayAutoCorrectOptions                    = $true    #   Show AutoCorrect Option buttons
$auc.correctInitialCaps                           = $false   #   Correct TWo INitial CApitals
$auc.correctSentenceCaps                          = $false   #   Capitalize first letter of sentences
$auc.correctTableCells                            = $false   #   Capitalize first letter of table cells
$auc.correctDays                                  = $true    #   Capitalize names of days
$auc.correctCapsLock                              = $false   #   Correct accidental usage of cAPS LOCK key
                                                             #
                                                             #
                                                             #   Replace as you type
                                                             #   --------------------------------------------------
$opt.autoFormatAsYouTypeReplaceQuotes             = $true    #  "Straight quotes" with “smart quotes“
$opt.autoFormatAsYouTypeReplaceFractions          = $true    #   Fractions (1/2) with fraction character ½
$opt.autoFormatAsYouTypeReplacePlainTextEmphasis  = $true    #  *Bold* and _italic_ with real formatting
$opt.autoFormatAsYouTypeReplaceHyperlinks         = $false   #   Internet and network paths with hyperlinks
$opt.autoFormatAsYouTypeReplaceOrdinals           = $true    #   Ordinals (1st) with superscript
$opt.autoFormatAsYouTypeReplaceSymbols            = $true    #   Hyphens (--) with dash
                                                             #
                                                             #
                                                             #   Apply as you type
                                                             #   ---------------------------------------------------
$opt.autoFormatAsYouTypeApplyBulletedLists        = $false   #   Automatic bulleted list
$opt.autoFormatAsYouTypeApplyBorders              = $false   #   Border lines
$opt.autoFormatAsYouTypeApplyHeadings             = $false   #   Built-in Heading styles
$opt.autoFormatAsYouTypeApplyNumberedLists        = $false   #   Automatic numbered lists
$opt.autoFormatAsYouTypeApplyTables               = $false   #   Tables
                                                             #
                                                             #
                                                             #   Automatically as you type
                                                             #   ----------------------------------------------------
$opt.autoFormatAsYouTypeFormatListItemBeginning   = $false   #   Format beginning of list item like the one before it
$opt.tabIndentKey                                 = $false   #   Set left- and first-indent with tabs and backspaces
$opt.autoFormatAsYouTypeDefineStyles              = $false   #   Define styles based on your formatting
                                                             #
                                                             #
                                                             #
                                                             #   W h e n   c o r r e c t i n g   s p e l l i n g
                                                             #   ===============================================
                                                             #
$opt.ignoreUppercase                              = $false   #   Ignore words in UPPERCASE
$opt.ignoreMixedDigits                            = $false   #   Ignore words that contain numbers
$opt.ignoreInternetAndFileAddresses               = $false   #   Ignore Internet and file addresses
$opt.repeatWord                                   = $true    #   Flag repeated words
$opt.allowAccentedUppercase                       = $false   #   Enforce accented uppercase in French
$opt.suggestFromMainDictionaryOnly                = $false   #   Suggest from main dictionary only
                                                             #
                                                             #
                                                             #
                                                             #   W h e n   c o r r e c t i n g   s p e l l i n g   a n d   g r a m m a r
                                                             #   =======================================================================
$opt.checkSpellingAsYouType                       = $true    #   Check-spelling as you type
$opt.checkGrammarAsYouType                        = $true    #   Mark grammar errors as you type
$opt.contextualSpeller                            = $true    #   Frequently confused words
$opt.checkGrammarWithSpelling                     = $true    #   Check grammar with spelling
#opt.????                                         = $false   #   Show readability statistics

$wrd.quit()
