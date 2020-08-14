using System;
using System.Collections.Generic;
using System.Text;
using b2xtranslator.Spreadsheet.XlsFileFormat.Ptg;
using b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using b2xtranslator.xls.XlsFileFormat.Structures;

namespace Macrome.Emulation
{
    public class Emulator
    {
        private WorkbookStream wbs;

        public Emulator(WorkbookStream wbs)
        {
            this.wbs = wbs;
        }

        public Lbl ResolveName(PtgName ptgName)
        {
            throw new NotImplementedException();
        }

        public Lbl ResolveName(PtgNameX ptgNameX)
        {
            throw new NotImplementedException();
        }


        public object ResolveRef(PtgRef ptgRef)
        {
            throw new NotImplementedException();
        }

        public List<Cell> GetAutoStartCells()
        {
            List<Cell> startCells = new List<Cell>();
            List<Lbl> autoOpenLabels = wbs.GetAutoOpenLabels();
            foreach (var autoOpenLbl in autoOpenLabels)
            {
                AbstractPtg peekPtg = autoOpenLbl.rgce.Peek();
                switch (peekPtg)
                {
                    //References a cell
                    case PtgRef3d ptgRef3d:
                        startCells.Add(new Cell(ptgRef3d.rw, ptgRef3d.col));
                        break;
                    //References a different label
                    case PtgName ptgName:
                        Lbl referencedLabel = ResolveName(ptgName);
                        while (referencedLabel.rgce.Peek() is PtgName)
                        {
                            referencedLabel = ResolveName(referencedLabel.rgce.Peek() as PtgName);
                        }
                        PtgRef3d ref3d = referencedLabel.rgce.Peek() as PtgRef3d;
                        startCells.Add(new Cell(ref3d.rw, ref3d.col));
                        break;
                    default:
                        throw new NotImplementedException();
                }

            }

            return startCells;
        }
    }
}
