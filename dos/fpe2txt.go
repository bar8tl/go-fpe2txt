// fpe2txt.go [2015-05-12 BAR8TL]
// Download Export Proforma Invoice to NAD .txt format - DOS mode
package main

import rb "bar8tl/p/fpe2txt"
import "fmt"
import "os"
import "regexp"

func main() {
  rb.PrepEnvironment()
  if _, err := os.Stat(os.Args[1]); err != nil {
    fmt.Printf("File %s does not extist, process stopped. Rectify.\n",
      os.Args[1])
    return
  }
  re := regexp.MustCompile("[.]([^.]+)$")
  txtnam := re.ReplaceAllString(os.Args[1], ".txt")
  if err := rb.Fpe2txt(os.Args[1], txtnam); err != nil {
    fmt.Printf(err.Error())
  } else {
    fmt.Printf("Success: Invoice downloaded successfully to format NAD .txt\n")
    fmt.Printf("Find result in file %s\n", txtnam)
  }
}
