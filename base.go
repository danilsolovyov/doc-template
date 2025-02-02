package docTemp

import (
	"bytes"
	"errors"
	"github.com/danilsolovyov/doc-template/docx"
	"log"
	"path/filepath"
	"strings"
	"text/template"
)

// Document interface is a combintation of methods use for generic data files
type Document interface {
	ReadFile(string) error
	UpdateContent(string)
	GetContent() string
	WriteToFile(string, string) error
	GetAsBytes(string) ([]byte, error)
	Close() error
}

// DocTemplate struct combines data and methods from both the Document interface
// and golang's templating library
type DocTemplate struct {
	Template *template.Template
	Document Document
}

// GetTemplate uses the file extension to determine the correct document struct to use
func GetTemplate(filePath string) (*DocTemplate, error) {
	var document Document
	switch filepath.Ext(filePath) {
	case ".docx", ".templ":
		document = new(docx.Docx)
	default:
		return nil, errors.New("Unsupported Document Type")
	}
	err := document.ReadFile(filePath)
	if err != nil {
		return nil, err
	}
	return &DocTemplate{Document: document, Template: template.New("docTemp")}, nil
}

func GetBytesTemplate(file []byte) (*DocTemplate, error) {
	document := new(docx.Docx)
	err := document.ReadBytes(file)
	if err != nil {
		return nil, err
	}
	return &DocTemplate{Document: document, Template: template.New("docTemp")}, nil
}

// Execute func runs the template and sends the output to the export path
func (docTemplate *DocTemplate) Execute(exportPath string, data interface{}) error {
	buf := new(bytes.Buffer)
	err := docTemplate.Template.Execute(buf, data)

	if err != nil {
		log.Println(err)
		return err
	}
	err = docTemplate.Document.WriteToFile(exportPath, strings.Replace(buf.String(), "<no value>", "", -1))
	return err
}

func (docTemplate *DocTemplate) ExecuteBytes(data interface{}) ([]byte, error) {
	buf := new(bytes.Buffer)
	err := docTemplate.Template.Execute(buf, data)

	if err != nil {
		log.Println(err)
		return nil, err
	}
	result, err := docTemplate.Document.GetAsBytes(strings.Replace(buf.String(), "<no value>", "", -1))

	if err != nil {
		log.Println(err)
		return nil, err
	}

	return result, err
}

// AddFunctions adds functions to the template
func (docTemplate *DocTemplate) AddFunctions(funcMap template.FuncMap) {
	docTemplate.Template = docTemplate.Template.Funcs(funcMap)
}

// Parse parses the template
func (docTemplate *DocTemplate) Parse() {
	temp, err := docTemplate.Template.Parse(docTemplate.Document.GetContent())
	if err != nil {
		log.Println(err)
	} else {
		docTemplate.Template = temp
	}
}
