package alignment

import (
	"testing"
	"unsafe"

	"github.com/Puzzanis/goxlsx.git/internal/xmltree"
)

func TestStructAlignment(t *testing.T) {
	t.Logf("XMLRow size: %d bytes", unsafe.Sizeof(xmltree.XMLRow{}))
	t.Logf("XMLCell size: %d bytes", unsafe.Sizeof(xmltree.XMLCell{}))
	t.Logf("InlineString size: %d bytes", unsafe.Sizeof(xmltree.InlineString{})) // Проверка смещения указателя в ячейке
	var c xmltree.XMLCell
	t.Logf("Offset of InlineString in XMLCell: %d", unsafe.Offsetof(c.InlineString))

	var s xmltree.SI
	t.Logf("Offset of Runs in SI: %d", unsafe.Offsetof(s.Runs))
}
