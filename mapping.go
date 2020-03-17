package main

type StoreSome struct {
	NameMedia	string
	Count		int
} 

type MapData struct {
	Year	int
	Data 	[]map[string]int
}

type Compare struct {
	Year	int
	Data 	[]map[string]int
}


func (s *MapData) ManipulateMM(v StoreSome)  {

	vr := map[string]int{v.NameMedia: v.Count}
	s.Data = append(s.Data, vr)
}

func (b *Compare) AddCompare(l StoreSome) {
	v := map[string]int{l.NameMedia: l.Count}
	b.Data = append(b.Data, v)
}


