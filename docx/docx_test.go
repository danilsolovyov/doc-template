package docx

import (
	"io/ioutil"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/stretchr/testify/assert"
)

type docxTest struct {
	fixture string
	content string
	err     error
}

var readDocTests = []docxTest{
	//  Check that reading a document works
	{fixture: "fixtures/test.docx", content: "This is a test document", err: nil},
}

func TestReadFile(t *testing.T) {
	for _, example := range readDocTests {
		actualDoc := new(Docx)
		actualErr := actualDoc.ReadFile(example.fixture)
		assert.Equal(t, example.err, actualErr)
		if actualErr == nil {
			assert.Contains(t, actualDoc.content, example.content)
		}
	}
}

func TestReadBytesFile(t *testing.T) {
	for _, example := range readDocTests {
		actualDoc := new(Docx)
		readBytes, actualErr := os.ReadFile(example.fixture)
		assert.Equal(t, example.err, actualErr)
		actualErr = actualDoc.ReadBytes(readBytes)

		assert.Equal(t, example.err, actualErr)
		if actualErr == nil {
			assert.Contains(t, actualDoc.content, example.content)
		}
	}
}

var writeDocTests = []docxTest{
	//  Check that writing a document works
	{fixture: "fixtures/test.docx", content: "This is an addition", err: nil},
}

func TestWriteToFile(t *testing.T) {
	for _, example := range writeDocTests {
		exportTempDir, _ := ioutil.TempDir("", "exports")
		// Overwrite content
		actualDoc := new(Docx)
		actualDoc.ReadFile(example.fixture)
		currentContent := actualDoc.GetContent()
		actualDoc.UpdateContent(strings.Replace(currentContent, "This is a test document", example.content, -1))
		newFilePath := filepath.Join(exportTempDir, "test.docx")
		actualDoc.WriteToFile(newFilePath, actualDoc.GetContent())
		// Check content
		newActualDoc := new(Docx)
		newActualDoc.ReadFile(newFilePath)
		assert.Contains(t, newActualDoc.GetContent(), example.content)
		os.RemoveAll(exportTempDir)
	}

}

func TestWriteToBytesFile(t *testing.T) {
	for _, example := range writeDocTests {
		exportTempDir, _ := ioutil.TempDir("", "exports")
		// Overwrite content
		actualDoc := new(Docx)
		readBytes, _ := os.ReadFile(example.fixture)
		actualDoc.ReadBytes(readBytes)
		currentContent := actualDoc.GetContent()
		actualDoc.UpdateContent(strings.Replace(currentContent, "This is a test document", example.content, -1))
		newFilePath := filepath.Join(exportTempDir, "test.docx")
		bytesData, actualErr := actualDoc.GetAsBytes(actualDoc.GetContent())
		assert.Equal(t, example.err, actualErr)
		actualErr = os.WriteFile(newFilePath, bytesData, os.ModePerm)
		assert.Equal(t, example.err, actualErr)
		// Check content
		newActualDoc := new(Docx)
		actualErr = newActualDoc.ReadFile(newFilePath)
		assert.Equal(t, example.err, actualErr)
		assert.Contains(t, newActualDoc.GetContent(), example.content)
		os.RemoveAll(exportTempDir)
	}

}
