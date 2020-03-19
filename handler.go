package main

import (
	"fmt"
	"log"
	"strconv"
	"strings"

	"github.com/360EntSecGroup-Skylar/excelize"
)

func ReadXls(fileName string, dataMap interface{}) ([]interface{}, error)  {
	var store []interface{}

	r, err := excelize.OpenFile(fileName)
	if err != nil {
		return nil, err
	}

	for i := 2011; i < 2020; i++ {
		if c, ok := dataMap.(MapData); ok {
			c.Year = i
			k := strconv.Itoa(i)
			rows := r.GetRows(k)
			for n, x := range rows {
				if n != 0 {
					count, _ := strconv.Atoi(x[1])
					mediaMe := x[0]
					var mn StoreSome
					if strings.Contains(mediaMe, "+") {
						// a := mediaMe
						c := strings.ReplaceAll(mediaMe, "+", "")
						if strings.Contains(c, "_") {
							a := strings.ReplaceAll(c, "_", "")
							b := strings.ToLower(a)
							mn.NameMedia = b
						}else {
							b := strings.ToLower(c)
							mn.NameMedia = b
						}
						mn.Count = count
					}
					if strings.Contains(mediaMe, " ") {
						// a := mediaMe
						a := strings.ReplaceAll(mediaMe, " ", "")
						if strings.Contains(a, "_") {
							b := strings.ToLower(strings.ReplaceAll(a, "_", ""))
							mn.NameMedia = b
						}else {
							b := strings.ToLower(a)
							mn.NameMedia = b
						}
						mn.Count = count
					}
					if !strings.Contains(mediaMe, "+") && !strings.Contains(mediaMe, " ") {
						a := mediaMe
						if strings.Contains(a, "_") {
							b := strings.ToLower(strings.ReplaceAll(a, "_", ""))
							mn.NameMedia = b
						}else {
							b := strings.ToLower(a)
							mn.NameMedia = b
						}
						mn.Count = count
					}
					c.ManipulateMM(mn)
				}			
			}
			store = append(store, c)
		}
	}
			
	return store, nil
}

func CompareVal(x *[]interface{}, y *[]interface{}) ([]interface{}, error) {	
	var storeData []interface{}
	for _, g := range *x {
		if s, ok := g.(MapData); ok {
			tmp := make(map[string][]map[string][]string, 0)
			for _, t := range s.Data {
				b, err := comMe(s.Year, t, y)
				if err != nil {
					return nil, err
				}
				yM := strconv.Itoa(s.Year)
				tmp[yM] = append(tmp[yM], b)
			}
			storeData = append(storeData, tmp)
		}
	}
	return storeData, nil
}


func comMe(year int, x map[string]int, k *[]interface{}) (map[string][]string, error) {
	res := map[string][]string{}
	for v, w := range x {
		for _, r := range *k {
			if s, ok := r.(MapData); ok {
				if year == s.Year {
					for _, b := range s.Data {
						store := make([]string, 0)
						if val, ok := b[v]; ok {
							st := val - w
							sAf := strconv.Itoa(st)
							valAf := strconv.Itoa(val)
							wAf := strconv.Itoa(w)
							store = append(store, wAf, valAf, sAf)
							res[v] = store
							return res, nil
						}else {
							wAf := strconv.Itoa(w)
							store = append(store, wAf, "NOTFOUND", "")
							res[v] = store
						}
					}
				}
			}
		}
	}

	return res, nil
}

func makeXls(x *[]interface{}, nameFile string) error  {
	xls := excelize.NewFile()
	for _, v := range *x {
		
		for x, y := range v.(map[string][]map[string][]string) {
			sheetName := x
			index := xls.NewSheet(sheetName)
			xls.SetCellValue(sheetName, "A1", "Source")
			xls.SetCellValue(sheetName, "B1", "Jumlah MMC")
			xls.SetCellValue(sheetName, "C1", "Jumlah Elastic IMA")
			xls.SetCellValue(sheetName, "D1", "Comparasi")

			for m, n := range y {
				for sm, sn := range n {
					xls.SetCellValue(sheetName, fmt.Sprintf("A%d",m+2), sm)
					if !ValInSlice(sn, "NOTFOUND") {
						for i, valM := range sn { 
							valMn, _ := strconv.Atoi(valM)
							switch	i {
							case 0:
								xls.SetCellValue(sheetName, fmt.Sprintf("B%d",m+2), valMn)
							case 1:
								xls.SetCellValue(sheetName, fmt.Sprintf("C%d",m+2), valMn)
							case 2:
								xls.SetCellValue(sheetName, fmt.Sprintf("D%d",m+2), valMn)
							}
						}
					}else {
						for i, valM := range sn { 
							switch	i {
							case 0:
								valMn, _ := strconv.Atoi(valM)
								xls.SetCellValue(sheetName, fmt.Sprintf("B%d",m+2), valMn)
							case 1:
								xls.SetCellValue(sheetName, fmt.Sprintf("C%d",m+2), valM)
							case 2:
								xls.SetCellValue(sheetName, fmt.Sprintf("D%d",m+2), valM)
							}
						}
					}
				}
			}
			xls.SetActiveSheet(index)
			log.Printf("%s.xlsx | sukses created sheet %s", nameFile, sheetName)
		}
	}
	
	err := xls.SaveAs(fmt.Sprintf("./%s.xlsx", nameFile))
	if err != nil {
		return nil
	}

	return nil
}

func ValInSlice(s []string, e string) bool  {
	for _, a := range s {
		if a == e {
			return true
		}
	}
	return false
}