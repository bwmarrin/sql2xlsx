// A simple example of quickly converting SQL result into an Excel file.
package main

import (
	"database/sql"
	"flag"
	"fmt"
	"io/ioutil"
	"log"
	"os"
	"time"

	_ "github.com/denisenkom/go-mssqldb"
	"github.com/tealeg/xlsx"
)

// Variables used with command line arguments
var (
	host string
	user string
	pass string
	sqlf string
	outf string
)

// Parse command line arguments
func init() {

	flag.StringVar(&host, "h", "", "SQL Server hostname or IP")
	flag.StringVar(&user, "u", "", "User ID")
	flag.StringVar(&pass, "p", "", "Password")
	flag.StringVar(&sqlf, "s", "", "SQL Query filename")
	flag.StringVar(&outf, "o", "", "Output filename")

	if len(os.Args) < 5 {
		flag.Usage()
		os.Exit(1)
	}

	flag.Parse()

}

func main() {

	// Connect to SQL Server
	dsn := fmt.Sprintf("server=%s;user id=%s;password=%s;encrypt=disable", host, user, pass)
	conn, err := sql.Open("mssql", dsn)
	if err != nil {
		log.Fatalf("error opening connection to server, %s\n", err)
	}
	defer conn.Close()

	// Verify connection to SQL Server
	err = conn.Ping()
	if err != nil {
		log.Fatalf("error verifying connection to server, %s\n", err)
		return
	}

	// Load SQL Query file
	query, err := ioutil.ReadFile(sqlf)
	if err != nil {
		log.Fatalf("error opening SQL Query file %s, %s\n", sqlf, err)
	}

	// Query the database
	rows, err := conn.Query(string(query))
	if err != nil {
		log.Fatalf("error running query, %s\n", err)
	}
	defer rows.Close()

	err = generateXLSXFromRows(rows, outf)
	if err != nil {
		log.Fatal(err)
	}

}

func generateXLSXFromRows(rows *sql.Rows, outf string) error {

	var err error

	// Get column names from query result
	colNames, err := rows.Columns()
	if err != nil {
		return fmt.Errorf("error fetching column names, %s\n", err)
	}
	length := len(colNames)

	// Create a interface slice filled with pointers to interface{}'s
	pointers := make([]interface{}, length)
	container := make([]interface{}, length)
	for i := range pointers {
		pointers[i] = &container[i]
	}

	// Create output xlsx workbook
	xfile := xlsx.NewFile()
	xsheet, err := xfile.AddSheet("Sheet1")
	if err != nil {
		return fmt.Errorf("error adding sheet to xlsx file, %s\n", err)
	}

	// Write Headers to 1st row
	xrow := xsheet.AddRow()
	xrow.WriteSlice(&colNames, -1)

	// Process sql rows
	for rows.Next() {

		// Scan the sql rows into the interface{} slice
		err = rows.Scan(pointers...)
		if err != nil {
			return fmt.Errorf("error scanning sql row, %s\n", err)
		}

		xrow = xsheet.AddRow()

		// Here we range over our container and look at each column
		// and set some different options depending on the column type.
		for _, v := range container {
			xcell := xrow.AddCell()
			switch v := v.(type) {
			case string:
				xcell.SetString(v)
			case []byte:
				xcell.SetString(string(v))
			case int64:
				xcell.SetInt64(v)
			case float64:
				xcell.SetFloat(v)
			case bool:
				xcell.SetBool(v)
			case time.Time:
				xcell.SetDateTime(v)
			default:
				xcell.SetValue(v)
			}

		}

	}

	// Save the excel file to the provided output file
	err = xfile.Save(outf)
	if err != nil {
		return fmt.Errorf("error writing to output file %s, %s\n", outf, err)
	}

	return nil
}
