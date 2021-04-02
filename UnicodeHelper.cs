using System;
using System.Collections.Generic;
using System.Text;

namespace Macrome
{
    public static class UnicodeHelper
    {
        //This class provides examples of how to abuse a few comparison "features" in Excel
        //1. Null bytes are ignored at the beginning and start of a label.
        //2. Comparisons are not case sensitive, A vs a or Ḁ vs ḁ
        //3. Unicode strings can be "decomposed" - ex: Ḁ (U+1E00) can become A (U+0041) - ◌̥ (U+0325)
        //4. The Combining Grapheme Joiner (U+034F) unicode symbol is ignored at any location in the string in SET.NAME functions
        //5. There a number of other "ignored" Unicode characters like \u200c, \u2060, and \uffef which can be used without impacting name usage

        //We can inject null bytes at the beginning and end of the string here without issue
        //We can also inject "ignored" unicode characters like non-breaking non-whitespace characters (\uffef, \u2060)
        //We can also inject invalid unicode characters like \udddd
        public const string UnicodeArgumentLabel = "\u0000v\u2060\u0041\u0325\u2060r";
        public const string UnicodeArgumentLabel2 = "\u0000v\u2060\u0041\u0325\u2060r2";

        //We can inject \u034f after any letter an arbitrary number of times since Excel will ignore it
        //https://en.wikipedia.org/wiki/Combining_Grapheme_Joiner - U+034F

        //Zero-Width Whitespace characters that are ignored for comparisons include \u2060, \u200c, and \u200d
        //\uffef is also ignored, though it will create a visible space
        //Invalid unicode values like \udddd can also be used here, though they will create a visible artifact of use
        public const string VarName = "v\u2060\u200c\u200d\u1e01\u034fr";
        public const string VarName2 = "v\u2060\u200c\u200d\u1e01\u034fr2";


        //\u180e and \u200b are considered zero width whitespace characters that can break a string
        //Decoy var names can be started (or injected) with a unicode whitespace character that breaks comparisons but looks identical
        //This makes it very frustrating to compare strings visually
        public const string DecoyVarName = "v\u200b\u200b\u200b\u1e01\u180er";

        public const string CharFuncArgument1Label = "cfv1";
        public const string FormulaFuncArgument1Label = "ffv1";
        public const string FormulaFuncArgument2Label = "ffv2";
        public const string FormulaEvalArgument1Label = "fev1";
    }
}
