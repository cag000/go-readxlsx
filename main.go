package main

import "log"

const (
	MMCLOUD = `data news MMCloud(1).xlsx`
	ELASTICNEW = `elastic_180_215(NEW).xlsx`
	ELASTICOLD = `elastic_180_188(OLD)_online-news-isa-*.xlsx`
	ELASTICOLDV2 = `Elastic180188_V2.xlsx`
)


func main()  {
	var (
		mmCloud MapData
		elasticNew MapData
		elasticOld MapData
		elasticOldVT MapData
	)

	mmC, err := ReadXls(MMCLOUD, mmCloud)
	if err != nil {
		log.Fatal(err.Error())
	}

	esN, err := ReadXls(ELASTICNEW, elasticNew)
	if err != nil {
		log.Fatal(err.Error())
	}

	esOO, err := ReadXls(ELASTICOLD, elasticOld)
	if err != nil {
		log.Fatal(err.Error())
	}
	// log.Print(mmC...)

	esOOVT, err := ReadXls(ELASTICOLDV2, elasticOldVT)
	if err != nil {
		log.Fatal(err.Error())
	}
	
	//mmcloud x elastic new
	mmcXesN, err := CompareVal(&mmC, &esN)
	if err != nil {
		log.Fatal(err.Error())
	}
	// log.Print(mmcXesN...)

	//mmcloud x elastic old index online-news-isa-*
	mmcXesOO, err := CompareVal(&mmC, &esOO)
	if err != nil {
		log.Fatal(err.Error())
	}
	// log.Print(mmcXesOO...)

	//mmcloud x elastic old index online-news-isa-v2-*
	mmcXesOOVT, err := CompareVal(&mmC, &esOOVT)
	if err != nil {
		log.Fatal(err.Error())
	}

	//make XLSX mmcloud x elastic new
	err = makeXls(&mmcXesN, "MMCxESNEW")
	if err != nil {
		log.Fatal(err.Error())
	}

	//make XLSX mmcloud x elastic old index online-news-isa-*
	err = makeXls(&mmcXesOO, "MMCxESOLD_online-news-isa")
	if err != nil {
		log.Fatal(err.Error())
	}

	//make XLSX mmcloud x elastic old index online-news-isa-v2-*
	err = makeXls(&mmcXesOOVT, "MMCxESOLD_online-news-isa-v2")
	if err != nil {
		log.Fatal(err.Error())
	}


	// log.Print(mmcXesN)
	// log.Print(mmcXesOO)

}