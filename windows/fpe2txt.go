// fpe2txt.go [2015-05-12 BAR8TL]
// Download Export Proforma Invoice to NAD .txt format - Windows mode
package main

import rb "bar8tl/p/fpe2txt"
import "github.com/lxn/walk"
import . "github.com/lxn/walk/declarative"
import "log"
import "regexp"

type MyMainWindow struct {
  *walk.MainWindow
  xlsnam *walk.LineEdit
}

func main() {
  rb.PrepEnvironment()
  mw := new(MyMainWindow)
  if err := ProcMainDialog(*mw); err != nil {
    log.Fatal(err)
  }
}

func (mw *MyMainWindow) downloadInvoice() {
  if len(mw.xlsnam.Text()) == 0 {
    walk.MsgBox(mw, "Error", "File 'Export Proforma Invoice' not specified.\n" +
      "Specify a valid XLSX file.", walk.MsgBoxIconError)
    return
  }
  re := regexp.MustCompile("[.]([^.]+)$")
  txtnam := re.ReplaceAllString(mw.xlsnam.Text(), ".txt")
  if err := rb.Fpe2txt(mw.xlsnam.Text(), txtnam); err != nil {
    walk.MsgBox(mw, "Error", err.Error(), walk.MsgBoxIconError)
  } else {
    walk.MsgBox(mw, "Success", "Invoice has been downloaded successfully in " +
      "TXT format.\nFind output in file "+txtnam, walk.MsgBoxIconInformation)
  }
}

func ProcMainDialog(mw MyMainWindow) (err error) {
  _, err = (MainWindow{
    AssignTo: &mw.MainWindow,
    Title:    "Download Export Proforma Invoice to NAD .txt format",
    MinSize:  Size{650, 100},
    Layout:   VBox{},
    Children: []Widget{
      Composite{
        Layout: Grid{Columns: 3},
        Children: []Widget{
          Label{
            Text: "Export Proforma Invoice:",
          },
          LineEdit{
            AssignTo: &(mw.xlsnam),
            ReadOnly: true,
          },
          PushButton{
            Text:      "Browse...",
            OnClicked: mw.openInvoice_Triggered,
          },
        },
      },
      Composite{
        Layout: HBox{},
        Children: []Widget{
          HSpacer{},
          PushButton{
            Text:      "Download",
            OnClicked: mw.downloadInvoice,
          },
          PushButton{
            Text:      "Close",
            OnClicked: func() { mw.Close() },
          },
        },
      },
    },
  }.Run())
  return
}

func (mw *MyMainWindow) openInvoice_Triggered() {
  if err := mw.openXls(); err != nil {
    log.Print(err)
  }
}

func (mw *MyMainWindow) openXls() error {
  dlg := new(walk.FileDialog)
  dlg.Filter = "XLSX Files (*.xlsx)|*.xlsx"
  dlg.Title = "Select an Export Proforma Invoice"
  if ok, err := dlg.ShowOpen(mw); err != nil {
    return err
  } else if !ok {
    return nil
  }
  mw.xlsnam.SetText(dlg.FilePath)
  return nil
}