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
						a := strings.ReplaceAll(mediaMe, "+", "")
						b := strings.ToLower(a)
						mn.NameMedia = b
						mn.Count = count
					}
					if strings.Contains(mediaMe, " ") {
						// a := mediaMe
						a := strings.ReplaceAll(mediaMe, " ", "")
						b := strings.ToLower(a)
						mn.NameMedia = b
						mn.Count = count
					}
					if !strings.Contains(mediaMe, "+") && !strings.Contains(mediaMe, " ") {
						a := mediaMe
						b := strings.ToLower(a)
						mn.NameMedia = b
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
	var storeData	[]interface{}
	for _, g := range *x {
		if s, ok := g.(MapData); ok {
			tmp := make(map[string][]map[string]string, 0)
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


func comMe(year int, x map[string]int, k *[]interface{}) (map[string]string, error) {
	res := map[string]string {}
	for v, w := range x {
		for _, r := range *k {
			if s, ok := r.(MapData); ok {
				if year == s.Year {
					for _, b := range s.Data {
						if val, ok := b[v]; ok {
							s := val - w
							res[v] = strconv.Itoa(s)
							return res, nil
						}else {
							res[v] = "Media_Not_Found"
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
		
		for x, y := range v.(map[string][]map[string]string) {
			sheetName := x
			index := xls.NewSheet(sheetName)
			xls.SetCellValue(sheetName, "A1", "Source")
			xls.SetCellValue(sheetName, "B1", "Comparasi")
			for m, n := range y {
				for k, l := range n {
					xls.SetCellValue(sheetName, fmt.Sprintf("A%d", m+2), k)
					if l != "Media_Not_Found" {
						l, _ := strconv.Atoi(l)
						xls.SetCellValue(sheetName, fmt.Sprintf("B%d", m+2), l)	
					}else {
						xls.SetCellValue(sheetName, fmt.Sprintf("B%d", m+2), l)	
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