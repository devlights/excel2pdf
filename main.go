package main

import (
	"flag"
	"fmt"
	"log/slog"
	"os"
	"path/filepath"
	"strings"

	"github.com/devlights/goxcel"
	"github.com/devlights/goxcel/constants"
)

type target struct {
	filePath string
	absPath  string
	pdfPath  string
	verbose  bool
}

func (me *target) abs() string {
	if me.absPath == "" {
		v, _ := filepath.Abs(me.filePath)
		me.absPath = v
	}

	if me.verbose {
		slog.Info("abs", "path", me.absPath)
	}

	return me.absPath
}

func (me *target) convert() string {
	if me.absPath == "" {
		me.abs()
	}

	me.pdfPath = me.absPath[:strings.Index(me.absPath, filepath.Ext(me.absPath))] + ".pdf"

	if me.verbose {
		slog.Info("convert", "abs", me.absPath, "pdf", me.pdfPath)
	}

	return me.pdfPath
}

func main() {
	var (
		verbose bool
	)

	flag.Usage = func() {
		fmt.Fprintln(os.Stderr, "usage: excel2pdf.exe (-v) excel-file-path")
		flag.PrintDefaults()
	}

	flag.BoolVar(&verbose, "v", false, "verbose log output")
	flag.Parse()

	if flag.NArg() < 1 {
		flag.Usage()
		return
	}

	if err := run(&target{filePath: flag.Arg(0), verbose: verbose}); err != nil {
		slog.Error(err.Error())
	}
}

func run(p *target) error {
	if p.verbose {
		slog.Info("start")
		defer slog.Info("done")
	}

	quitFn := goxcel.MustInitGoxcel()
	defer quitFn()

	g, goxcelReleaseFn := goxcel.MustNewGoxcel()
	defer goxcelReleaseFn()

	_ = g.Silent(false)

	wbs, err := g.Workbooks()
	if err != nil {
		return err
	}

	wb, wbReleaseFn, err := wbs.Open(p.abs())
	if err != nil {
		return err
	}
	defer wbReleaseFn()

	err = wb.ExportAsFixedFormat(constants.XlTypePDF, p.convert())
	if err != nil {
		return err
	}

	return nil
}
