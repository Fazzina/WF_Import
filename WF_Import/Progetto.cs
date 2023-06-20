using System;
using System.Collections.Generic;

namespace WF_Import
{
    public class Progetto
    {
        public string Anno { get; set; }
        public int nProtocollo { get; set; }
        public DateTime? dataInserimento { get; set; }
        public string Tipologia { get; set; }
        public string Stato { get; set; }
        public string Ambito { get; set; }
        public string Soggetti { get; set; }
        public string Titolo { get; set; }
        public DateTime? dataInizio { get; set; }
        public DateTime? dataFine { get; set; }
        public string Errori { get; set; }
    }

    class Config
    {
        public int FirstRow { get; set; }
        public List<FieldMap> fieldsMap { get; set; }
    }

    class FieldMap
    {
        public int columnID { get; set; }
        public string datatype { get; set; }
    }


}