#
#
#                                         Microsoft Word Proofing options:
#
#                                         Change how Word corrects and formats your text.
#
#

$wrd = new-object -com word.application
$wrd.visible = $true

$opt = $wrd.options


                                                             #   A u t o C o r r e c t   O p t i o n s
                                                             #   =====================================
                                                             #
                                                             #   Replace as you type
                                                             #   --------------------------------------------------
$opt.autoFormatAsYouTypeReplaceQuotes             = $false   #  "Straight quotes" with “smart quotes“
$opt.autoFormatAsYouTypeReplaceFractions          = $false   #   Fractions (1/2) with fraction character ½
$opt.autoFormatAsYouTypeReplacePlainTextEmphasis  = $false   #  *Bold* and _italic_ with real formatting
$opt.autoFormatAsYouTypeReplaceHyperlinks         = $false   #   Internet and network paths with hyperlinks
$opt.autoFormatAsYouTypeReplaceOrdinals           = $false   #   Ordinals (1st) with superscript
$opt.autoFormatAsYouTypeReplaceSymbols            = $false   #   Hyphens (--) with dash
                                                             #
                                                             #
                                                             #   Apply as you type
                                                             #   ---------------------------------------------------
$opt.autoFormatAsYouTypeApplyBulletedLists        = $false   #   Automatic bulleted list
$opt.autoFormatAsYouTypeApplyBorders              = $false   #   Border lines
$opt.autoFormatAsYouTypeApplyHeadings             = $false   #   Built-in Heading styles
$opt.autoFormatAsYouTypeApplyNumberedLists        = $false   #   Automatic numbered lists
$opt.autoFormatAsYouTypeApplyNumberedLists        = $false   #   Tables
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
$opt.repeatWord                                   = $false   #   Flag repeated words
$opt.allowAccentedUppercase                       = $false   #   Enforce accented uppercase in French
$opt.suggestFromMainDictionaryOnly                = $false   #   Suggest from main dictionary only
                                                             #
                                                             #
                                                             #
                                                             #   W h e n   c o r r e c t i n g   s p e l l i n g   a n d   g r a m m a r
                                                             #   =======================================================================
$opt.checkSpellingAsYouType                       = $false   #   Check-spelling as you type
$opt.checkGrammarAsYouType                        = $false   #   Mark grammar errors as you type
$opt.contextualSpeller                            = $false   #   Frequently confused words
$opt.checkGrammarWithSpelling                     = $true    #   Check grammar with spelling
#opt.????                                         = $false   #   Show readability statistics

$wrd.quit()
