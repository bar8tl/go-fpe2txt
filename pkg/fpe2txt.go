// fpe2txt.go [2015-05-12 BAR8TL]
// Download Export Proforma Invoice to NAD .txt format
package rb

import "code.google.com/p/gcfg"
import "errors"
import "fmt"
import "github.com/tealeg/xlsx"
import "os"

const PROFORMA_INVOICE   = "FACTURA PROFORMA DE EXPORTACION"
const PIPE               = "|"
const TILDE              = "~"
const AT                 = "@"
const ITEMS_TOP_LINE     = 20
const ITEMS_BOTTOM_LINE  = 43
const FIELDS_HEADER      = 15
const FIELDS_ITEMS       = 16
const FIELDS_INDICATORS  =  7
const FIELDS_PERMISSIONS =  6

type sSttg struct {
  Dtfmt string
}

type ml struct {
  seq int
  dsc string
  otp string
  lon int
  frm string
  idx int
  itp string
  mtd string
  row int
  col int
  dcn int
  scn string
}

type hl [FIELDS_HEADER]      string
type il [FIELDS_ITEMS]       string
type dl [FIELDS_INDICATORS]  string
type pl [FIELDS_PERMISSIONS] string

var S     sSttg
var hmap, imap []ml
var uom   map[string]string

func Fpe2txt(xlsnam, txtnam string) error {
  xlFile, err := xlsx.OpenFile(xlsnam)
  if err != nil {
    return err
  }
  sheet := xlFile.Sheets[0]
  if sheet.Cell(0, 0).String() != PROFORMA_INVOICE {
    return errors.New("Error: Document is not an Export Proforma Invoice.")
  }
  bldUomMap(); bldMapHdr(); bldMapItm()
  err = validateDocument(sheet)
  if err != nil {
    return err
  }
  f, err := os.Create(txtnam)
  if err != nil {
    return err
  }
  var hdr hl
  mapHeader(sheet, &hdr)
  var itm []il
  for rn := ITEMS_TOP_LINE; rn < ITEMS_BOTTOM_LINE; rn++ {
    if sheet.Cell(rn, 1).String() != "" {
      var it il
      mapItem(rn, sheet, &it)
      itm = append(itm, it)
    }
  }
  var ind []dl
  var per []pl
  o := bldStringOneInvoice(&hdr, itm, ind, per)
  _, err = f.WriteString(o)
  return err
}

func mapHeader(sheet *xlsx.Sheet, hdr *hl) {
  for i := 0; i < len(hmap); i++ {
    var st string
    if hmap[i].mtd == "const" {
      if hmap[i].otp == "d" {
        st = fmt.Sprintf(hmap[i].frm, hmap[i].dcn)
      } else if hmap[i].otp == "s" {
        st = fmt.Sprintf(hmap[i].frm, hmap[i].scn)
      }
    } else if hmap[i].mtd == "map" {
      if hmap[i].otp == "d" {
        dl, _ := sheet.Cell(hmap[i].row, hmap[i].col).Int()
        st = fmt.Sprintf(hmap[i].frm, dl)
      } else if hmap[i].otp == "f" {
        fl, _ := sheet.Cell(hmap[i].row, hmap[i].col).Float()
        st = fmt.Sprintf(hmap[i].frm, fl)
      } else {
        sl := sheet.Cell(hmap[i].row, hmap[i].col).String()
        if hmap[i].idx == 2 {
          st = formatDate(sl)
        } else {
          st = fmt.Sprintf(hmap[i].frm, sl)
        }
      }
    }
    hdr[hmap[i].idx] = st
  }
}

func mapItem(rn int, sheet *xlsx.Sheet, it *il) {
  for i := 0; i < len(imap); i++ {
    var st string

    if imap[i].mtd == "const" {
      if imap[i].otp == "d" {
        st = fmt.Sprintf(imap[i].frm, imap[i].dcn)
      } else if imap[i].otp == "s" {
        st = fmt.Sprintf(imap[i].frm, imap[i].scn)
      }
    } else if imap[i].mtd == "map" {
      if imap[i].otp == "d" {
        dl, _ := sheet.Cell(rn, imap[i].col).Int()
        st = fmt.Sprintf(imap[i].frm, dl)
      } else if imap[i].otp == "f" {
        fl, _ := sheet.Cell(rn, imap[i].col).Float()
        st = fmt.Sprintf(imap[i].frm, fl)
      } else {
        sl := sheet.Cell(rn, imap[i].col).String()
        if imap[i].idx == 4 {
          st = fmt.Sprintf(imap[i].frm, uom[sl])
        } else {
          st = fmt.Sprintf(imap[i].frm, sl)
        }
      }
    }
    it[imap[i].idx] = st
  }
}

func bldMapHdr() {
  hmap = append(hmap, ml{ 1, "Numero de factura",        "s",  50, "%s",     0,
    "f", "map",    7,  8, 0, ""   })
  hmap = append(hmap, ml{ 2, "Numero de pedido",         "s",  50, "%s",     1,
    "-", "empty",  0,  0, 0, ""   })
  hmap = append(hmap, ml{ 3, "Fecha de facturacion",     "s",  10, "%s",     2,
    "s", "map",    7,  5, 0, ""   })  // dd/mm/yyyy
  hmap = append(hmap, ml{ 4, "Pais de facturacion",      "s",   3, "%s",     3,
    "-", "const",  0,  0, 0, "MEX"})  // MEX
  hmap = append(hmap, ml{ 5, "Entidad de facturacion",   "s",   2, "%s",     4,
    "-", "const",  0,  0, 0, "EM" })  // EM
  hmap = append(hmap, ml{ 6, "Moneda",                   "s",   3, "%s",     5,
    "s", "map",   19, 10, 0, ""   })
  hmap = append(hmap, ml{ 7, "Termino",                  "s",   3, "%s",     6,
    "s", "map",   15, 10, 0, ""   })
  hmap = append(hmap, ml{ 8, "Valor moneda extranjera",  "f",  14, "%.8f",   7,
    "f", "map",   45, 10, 0, ""   })  // 14,8
  hmap = append(hmap, ml{ 9, "Valor comercial",          "f",  14, "%.8f",   8,
    "f", "map",   45, 10, 0, ""   })  // 14,8
  hmap = append(hmap, ml{10, "Flete",                    "d",  14, "%d",     9,
    "-", "const",  0,  0, 0, ""   })  // 14,8 const = 0
  hmap = append(hmap, ml{11, "Seguros",                  "d",  14, "%d",    10,
    "-", "const",  0,  0, 0, ""   })  // 14,8 const = 0
  hmap = append(hmap, ml{12, "Embalajes",                "d",  14, "%d",    11,
    "-", "const",  0,  0, 0, ""   })  // 14,8 const = 0
  hmap = append(hmap, ml{13, "Incrementables",           "d",  14, "%d",    12,
    "-", "const",  0,  0, 0, ""   })  // 14,8 const = 0
  hmap = append(hmap, ml{14, "Deducibles",               "d",  14, "%d",    13,
    "-", "const",  0,  0, 0, ""   })  // 14,8 const = 0
  hmap = append(hmap, ml{15, "Factor moneda extranjera", "f",  14, "%.8f",  14,
    "f", "map",   43,  1, 0, ""   })  // 14,8
}

func bldMapItm() {
  imap = append(imap, ml{16, "Numero de parte",          "s",  30, "%s",     0,
    "s", "map",    0,  1, 0, ""   })
  imap = append(imap, ml{17, "Descripcion en español",   "s", 255, "%s",     1,
    "-", "empty",  0,  0, 0, ""   })
  imap = append(imap, ml{18, "Descripcion en ingles",    "s", 255, "%s",     2,
    "s", "map",    0,  2, 0, ""   })
  imap = append(imap, ml{19, "Cantidad UMC",             "f",  14, "%.9f",   3,
    "f", "map",    0,  6, 0, ""   })  // 14,9
  imap = append(imap, ml{20, "UMC",                      "s",   2, "%s",     4,
    "s", "map",    0,  4, 0, ""   })
  imap = append(imap, ml{21, "Precio unitario",          "f",  14, "%.9f",   5,
    "f", "map",    0,  9, 0, ""   })  // 14,9
  imap = append(imap, ml{22, "Unidad peso unitario",     "f",   1, "%f",     6,
    "-", "empty",  0,  0, 0, ""   })
  imap = append(imap, ml{23, "Peso unitario",            "f",  14, "%.3f",   7,
    "-", "empty",  0,  0, 0, ""   })
  imap = append(imap, ml{24, "Fraccion",                 "d",   8, "%d",     8,
    "f", "map",    0,  5, 0, ""   })
  imap = append(imap, ml{25, "Cantidad UMT",             "d",   8, "%d",     9,
    "-", "const",  0,  0, 1, ""   })  // 14,9 const = 1
  imap = append(imap, ml{26, "Factor de ajuste",         "f",   3, "%.12f", 10,
    "-", "empty",  0,  0, 0, ""   })
  imap = append(imap, ml{27, "Pais origen",              "s",   3, "%s",    11,
    "s", "map",    0,  3, 0, ""   })
  imap = append(imap, ml{28, "Valor agregado",           "f",  14, "%.9f",  12,
    "f", "map",    0,  8, 0, ""   })  // 14,9
  imap = append(imap, ml{29, "Marca",                    "s",  70, "%s",    13,
    "-", "empty",  0,  0, 0, ""   })
  imap = append(imap, ml{30, "Modelo",                   "s",  70, "%s",    14,
    "-", "empty",  0,  0, 0, ""   })
  imap = append(imap, ml{31, "Serie",                    "s",  70, "%s",    15,
    "-", "empty",  0,  0, 0, ""   })
  }
func bldUomMap() {
  uom = make(map[string]string)
  uom["KGS" ] = "01"
  uom["GR"  ] = "02"
  uom["M"   ] = "03"
  uom["M2"  ] = "04"
  uom["M3"  ] = "05"
  uom["PZS" ] = "06"
  uom["CAB" ] = "07"
  uom["LT"  ] = "08"
  uom["PAR" ] = "09"
  uom["KW"  ] = "10"
  uom["MIL" ] = "11"
  uom["JGO" ] = "12"
  uom["KWH" ] = "13"
  uom["TON" ] = "14"
  uom["BAR" ] = "15"
  uom["GRN" ] = "16"
  uom["DEC" ] = "17"
  uom["CEN" ] = "18"
  uom["DOZ" ] = "19"
  uom["CAJA"] = "20"
  uom["BOT" ] = "21"
}

func bldStringOneInvoice(hdr *hl, itm []il, ind []dl, per []pl) string {
  var s string
  for i := 0; i < len(hdr); i++ {
    if i == 0 {
      s = hdr[i]
    } else {
      s += PIPE + hdr[i]
    }
  }
  for i := 0; i < len(itm); i++ {
    for j := 0; j < FIELDS_ITEMS; j++ {
      s += PIPE + itm[i][j]
    }
  }
  s += TILDE
  for i := 0; i < len(ind); i++ {
    for j := 0; j < FIELDS_INDICATORS; j++ {
      if i == 0 && j == 0 {
        s += ind[i][j]
      } else {
        s += PIPE + ind[i][j]
      }
    }
  }
  s += AT
  for i := 0; i < len(per); i++ {
    for j := 0; j < FIELDS_PERMISSIONS; j++ {
      if i == 0 && j == 0 {
        s += per[i][j]
      } else {
        s += PIPE + per[i][j]
      }
    }
  }
  return s
}

func validateDocument(sheet *xlsx.Sheet) error {
  ne := 1
  es := validateHeadr(sheet, &ne)
  es += validateItems(sheet, &ne)
  if es != "" {
    return errors.New("Errors:\n"+es)
  }
  return nil
}

func validateHeadr(sheet *xlsx.Sheet, ne *int) (es string) {
  for i := 0; i < len(hmap); i++ {
    if hmap[i].itp == "f" {
      _, err := sheet.Cell(hmap[i].row, hmap[i].col).Float()
      if err != nil {
        es += fmt.Sprintf("%d) Mandatory field \"%s\" is empty " +
          "or contains a non-numeric character\n", *ne, hmap[i].dsc)
        (*ne)++
      }
    } else if hmap[i].itp == "s" {
      sl := sheet.Cell(hmap[i].row, hmap[i].col).String()
      if sl == "" {
        es += fmt.Sprintf("%d) Mandatory field \"%s\" is empty\n",
          *ne, hmap[i].dsc)
        (*ne)++
      }
    }
  }
  return es
}

func validateItems(sheet *xlsx.Sheet, ne *int) (es string) {
  for rn := ITEMS_TOP_LINE; rn < ITEMS_BOTTOM_LINE; rn++ {
    if sheet.Cell(rn, 1).String() != "" {
      for i := 0; i < len(imap); i++ {
        if imap[i].itp == "f" {
          _, err := sheet.Cell(rn, imap[i].col).Float()
          if err != nil {
            es += fmt.Sprintf("%d) Line %d: Mandatory field \"%s\" is empty " +
              "or contains a non-numeric character\n", *ne, rn+1, imap[i].dsc)
            (*ne)++
          }
        } else if imap[i].itp == "s" {
          sl := sheet.Cell(rn, imap[i].col).String()
          if sl == "" {
            es += fmt.Sprintf("%d) Line %d: Mandatory field \"%s\" is empty\n",
              *ne, rn+1, imap[i].dsc)
            (*ne)++
          }
        }
      }
    }
  }
  return es
}

func formatDate(sl string) (st string) {
  switch S.Dtfmt {
    case "dd-mm-yy"  : st = fmt.Sprintf("%s/%s/20%s",sl[0: 2],sl[3:5],sl[6: 8])
    case "mm-dd-yy"  : st = fmt.Sprintf("%s/%s/20%s",sl[3: 5],sl[0:2],sl[6: 8])
    case "yy-mm-dd"  : st = fmt.Sprintf("%s/%s/20%s",sl[6: 8],sl[3:5],sl[0: 2])
    case "ddmmyy"    : st = fmt.Sprintf("%s/%s/20%s",sl[0: 2],sl[2:4],sl[4: 6])
    case "mmddyy"    : st = fmt.Sprintf("%s/%s/20%s",sl[2: 4],sl[0:2],sl[4: 6])
    case "yymmdd"    : st = fmt.Sprintf("%s/%s/20%s",sl[4: 6],sl[2:4],sl[0: 2])
    case "dd-mm-yyyy": st = fmt.Sprintf("%s/%s/%s",  sl[0: 2],sl[3:5],sl[6:10])
    case "mm-dd-yyyy": st = fmt.Sprintf("%s/%s/%s",  sl[3: 5],sl[0:2],sl[6:10])
    case "yyyy-mm-dd": st = fmt.Sprintf("%s/%s/%s",  sl[6:10],sl[3:5],sl[0: 2])
    case "ddmmyyyy"  : st = fmt.Sprintf("%s/%s/%s",  sl[0: 2],sl[2:4],sl[4: 8])
    case "mmddyyyy"  : st = fmt.Sprintf("%s/%s/%s",  sl[2: 4],sl[0:2],sl[4: 8])
    case "yyyymmdd"  : st = fmt.Sprintf("%s/%s/%s",  sl[4: 8],sl[2:4],sl[0: 2])
    case "", "na"    : st = fmt.Sprintf("%s", sl)
  }
  return st
}

func PrepEnvironment() {
  GetConfig("fpe2txt.gcfg")
}

func GetConfig(f string) {
  var Cfg struct {
    OwnSettings struct {
      DlDateFormat string
    }
  }
  gcfg.ReadFileInto(&Cfg, f)
  S.Dtfmt = Cfg.OwnSettings.DlDateFormat
}
