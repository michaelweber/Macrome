using System.Collections.Generic;
using b2xtranslator.CommonTranslatorLib;
using b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class ObjectsSequence : BiffRecordSequence, IVisitable
    {
        public MsoDrawingSelection MsoDrawingSelection;

        public List<DrawingsGroup> DrawingsGroup;

        public ObjectsSequence(IStreamReader reader)
            : base(reader)
        {
            // OBJECTS = *(MsoDrawing *Continue *(TEXTOBJECT / OBJ / CHART)) [MsoDrawingSelection]

            // *(MsoDrawing *Continue *(TEXTOBJECT / OBJ / CHART))
            this.DrawingsGroup = new List<DrawingsGroup>();
            while (BiffRecord.GetNextRecordType(reader) == RecordType.MsoDrawing)
            {
                this.DrawingsGroup.Add(new DrawingsGroup(reader));
            }

            // [MsoDrawingSelection]
            if (BiffRecord.GetNextRecordType(reader) == RecordType.MsoDrawingSelection)
            {
                this.MsoDrawingSelection = (MsoDrawingSelection)BiffRecord.ReadRecord(reader);
            }
        }

        #region IVisitable Members

        public void Convert<T>(T mapping)
        {
            ((IMapping<ObjectsSequence>)mapping).Apply(this);
        }

        #endregion
    }

    public class DrawingsGroup
    {
        public MsoDrawing MsoDrawing;

        public List<Continue> Continues;

        public List<ObjectGroup> Objects;

        public DrawingsGroup(IStreamReader reader)
        {
            // MsoDrawing *Continue *(TEXTOBJECT / OBJ / CHART)

            // MsoDrawing
            this.MsoDrawing = (MsoDrawing)BiffRecord.ReadRecord(reader);

            // *Continue
            this.Continues = new List<Continue>();
            while (BiffRecord.GetNextRecordType(reader) == RecordType.Continue)
            {
                this.Continues.Add((Continue)BiffRecord.ReadRecord(reader));
            }

            // *(TEXTOBJECT / OBJ / CHART)
            this.Objects = new List<ObjectGroup>();
            while (BiffRecord.GetNextRecordType(reader) == RecordType.Obj
                || BiffRecord.GetNextRecordType(reader) == RecordType.TxO
                || BiffRecord.GetNextRecordType(reader) == RecordType.BOF)
            {
                this.Objects.Add(new ObjectGroup(reader));
            }
        }
    }

    public class ObjectGroup
    {
        public TextObjectSequence TextObjectSequence;

        public Obj Obj;

        public ChartSheetSequence ChartSheetSequence;

        public ObjectGroup(IStreamReader reader)
        {
            if (BiffRecord.GetNextRecordType(reader) == RecordType.Obj)
            {
                this.Obj = (Obj)BiffRecord.ReadRecord(reader);
            }
            else if (BiffRecord.GetNextRecordType(reader) == RecordType.TxO)
            {
                this.TextObjectSequence = new TextObjectSequence(reader);
            }
            else if (BiffRecord.GetNextRecordType(reader) == RecordType.BOF)
            {
                this.ChartSheetSequence = new ChartSheetSequence(reader);
            }
        }
    }
}
