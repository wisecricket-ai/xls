package xls

type rowInfo struct {
	Index    uint16
	Fcell    uint16
	Lcell    uint16
	Height   uint16
	Notused  uint16
	Notused2 uint16
	Flags    uint32
}

// Row the data of one row
type Row struct {
	wb   *WorkBook
	info *rowInfo
	cols map[uint16]contentHandler
}

// Col Get the Nth Col from the Row, if has not, return nil.
// Suggest use Has function to test it.
func (r *Row) Col(i int) string {
	serial := uint16(i)
	if ch, ok := r.cols[serial]; ok {
		strs := ch.String(r.wb)
		return strs[0]
	} else {
		for _, v := range r.cols {
			if v.FirstCol() <= serial && v.LastCol() >= serial {
				strs := v.String(r.wb)
				index := serial - v.FirstCol()
				if int(index) < len(strs) {
					return strs[index]
				}
				// Only return a value if this is actually the correct column
				// Don't fall back to strs[0] as it may be from a different column
				if v.FirstCol() == serial {
					return strs[0]
				}
				// Return empty string if we can't find the exact column match
				return ""
			}
		}
	}
	return ""
}

// Raw Get the Nth Col from the Row without formatting, if has not, return nil.
func (r *Row) Raw(i int) string {
	serial := uint16(i)
	if ch, ok := r.cols[serial]; ok {
		return ch.RawValue(r.wb)
	} else {
		for _, v := range r.cols {
			if v.FirstCol() <= serial && v.LastCol() >= serial {
				// For multi-column cells, get the string array and return the specific position
				strs := v.String(r.wb)
				index := serial - v.FirstCol()
				if int(index) < len(strs) {
					// Return the value at this specific position, not the raw value of the whole cell
					return strs[index]
				}
				return ""
			}
		}
	}
	return ""
}

// ColExact Get the Nth Col from the Row, if has not, return nil.
// For merged cells value is returned for first cell only
func (r *Row) ColExact(i int) string {
	serial := uint16(i)
	if ch, ok := r.cols[serial]; ok {
		strs := ch.String(r.wb)
		return strs[0]
	}
	return ""
}

// LastCol Get the number of Last Col of the Row.
func (r *Row) LastCol() int {
	return int(r.info.Lcell)
}

// FirstCol Get the number of First Col of the Row.
func (r *Row) FirstCol() int {
	return int(r.info.Fcell)
}
