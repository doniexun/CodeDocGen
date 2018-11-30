package main

import (
	"fmt"
	"os"
	"io"
	"io/ioutil"
	"errors"
	"log"
	"strings"
	//"path/filepath"
	"bufio"
	"path"

	"baliance.com/gooxml/color"
	"baliance.com/gooxml/document"
	"baliance.com/gooxml/measurement"
	"baliance.com/gooxml/schema/soo/wml"
)

// 创建只有1*1大小的、带一级标题的表格，并将内容填入
func createDocWithOnelineTable(doc *document.Document, title, content, filename string) (err error){
	
	if doc == nil {
		return errors.New("doc isn't exist")
	}

	// 添加页码
	//ftr := doc.AddFooter()
	//para := ftr.AddParagraph()
	//para.Properties().AddTabStop(3*measurement.Inch, wml.ST_TabJcCenter, wml.ST_TabTlcNone)

	//run := para.AddRun()
	//run.AddTab()
	//run.AddFieldWithFormatting(document.FieldCurrentPage, "", false)	// 当前页
	//run.AddText(" of ")
	//run.AddFieldWithFormatting(document.FieldNumberOfPages, "", false)	// 总页数
	//doc.BodySection().SetFooter(ftr, wml.ST_HdrFtrDefault)
	
	// 设置页面边距
	// func (s Section) SetPageMargins(top, right, bottom, left, header, footer, gutter measurement.Distance)
	doc.BodySection().SetPageMargins(10.0, 10.0, 10.0, 35.0, 10.0, 10.0, 10.0)

	// 添加标题
	para := doc.AddParagraph()
	para.Properties().Spacing().SetLineSpacing(12*measurement.Point, wml.ST_LineSpacingRuleAuto)
	para.SetStyle("Heading1")
	run := para.AddRun()
	run.Properties().SetFontFamily("Nimbus Roman No9 L")
	run.Properties().SetSize(10)
	run.AddText(title)			// 填入标题

	// 添加表格
	table := doc.AddTable()
	// 设置表格宽度
	table.Properties().SetWidthPercent(100)
	// 设置表格边框
	borders := table.Properties().Borders()
	borders.SetAll(wml.ST_BorderSingle, color.Auto, 1*measurement.Point)

	// 添加一行表格
	row := table.AddRow()
	run = row.AddCell().AddParagraph().AddRun()
	run.Properties().SetFontFamily("Nimbus Roman No9 L")
	run.Properties().SetSize(10)
	// 往表格中填入内容
	run.AddText(content)
	// 添加一个空行用以中断连续的表格
	doc.AddParagraph() // break up the consecutive tables

	// 检查文档，若无错发生，保存文档
	if err := doc.Validate(); err != nil {
		log.Fatalf("error during validation: %s", err)
	}
	doc.SaveToFile(filename)

	return nil
}

// 获取文件名后缀
func fileExt(filename string)(ext string, err error){
	//return filepath.Ext(filename), nil
	return path.Ext(filename), nil
}


// 读取指定文件内容
func fileContent(filename string)(content string, err error){
	file, err := os.Open(filename)
	if err != nil {
		fmt.Println(err)
		return "", err
	}
	defer file.Close()

	fd, err := ioutil.ReadAll(file)
	if err != nil {
		fmt.Println(err)
		return "", err
	}

	return string(fd), nil
}

// 读取指定文件内容
// 读一行，添加一个换行符
func fileContent2(filename string)(string, error){
	file, err := os.Open(filename)
	if err != nil {
		fmt.Println(err)
		return "", err
	}
	defer file.Close()

	content := ""
	buf := bufio.NewReader(file)
	for {
		line, err := buf.ReadString('\n')
		//fmt.Println(line)
		if err != nil {
			if err == io.EOF {
				return content, nil
			}
			return content, nil
		}
		content += line
	}
	return content, nil
}

// 遍历目录
func recursionDir(dir string)(err error){
	
	file, err := os.Open(dir)
	if err != nil {
		fmt.Println(err)
		return err
	}
	defer file.Close()
	
	files, err := file.Readdir(-1)
	if err != nil {
		fmt.Println(err)
		return err
	}

	for _, f := range files {
		if f.IsDir() {
			fmt.Println(f.Name(), " is a Dir")
			recursionDir(dir + "/" + f.Name())
		} else {
			fmt.Println(f.Name(), " is a file")
		}
	}
	
	return nil
}

func recursionDir2(doc *document.Document, sourceCodeDir, docName string, exceptDirs, exts []string)(err error){
	
	file, err := os.Open(sourceCodeDir)
	if err != nil {
		fmt.Println(err)
		return err
	}
	defer file.Close()
	
	files, err := file.Readdir(-1)
	if err != nil {
		fmt.Println(err)
		return err
	}

	for _, f := range files {
		if f.IsDir() {
			//fmt.Println(f.Name(), " is a Dir")
			if !isExceptDir(f.Name(), exceptDirs) {
				recursionDir2(doc, sourceCodeDir + "/" + f.Name(), docName, exceptDirs, exts)
			}
		} else {
			//fmt.Println(f.Name(), " is a file")
			// 判断是否是所需文件
			if isNeededFile(f.Name(), exts) {
				// 获取文件内容
				content, _ := fileContent2(sourceCodeDir + "/" + f.Name())
				//fmt.Println(content)
				// 将文件内容写入word中
				if err := createDocWithOnelineTable(doc, sourceCodeDir + "/" + f.Name(), content, docName); err != nil {
					panic(-1)
				}
			}
		}
	}
	
	return nil
}

// 判断是否是所需读取的文件
func isNeededFile(filename string, exts []string) bool {
	fileExt, _ := fileExt(filename)
	for _, ext := range exts {
		if fileExt == ext {
			return true
		}
	}
	return false
}

// 判断是否是所需排除的文件夹
func isExceptDir(dirName string, exceptDirs []string) bool {
	for _, exceptDir := range exceptDirs {
		if exceptDir == dirName {
			return true
		}
	}
	return false
}

// 从路径名中获取工程名
// 路径名示例："/home/windness/workspace/go/src/bookingSystem/"
// @TODO 暂时无法处理：1）H:/workspace/golang/booking   2）H:\workspace\golang\booking
func projectName(sourceCodeDir string) string {
	temp := ""
	pos := strings.LastIndex(sourceCodeDir, "/")
	if pos == len(sourceCodeDir) - 1 {	// 若路径以"/"结尾，则需要再次获取路径中的最后一个单词
		temp = sourceCodeDir[0 : len(sourceCodeDir) -1]
		pos = strings.LastIndex(temp, "/")
	}
	return temp[pos + 1:len(temp)] 
}

// 遍历指定目录，将指定后缀名文件中的文件写入word中
func generateCodeFile(sourceCodeDir string, exceptDirs []string, exts []string) error {

	// 创建DOC对象
	doc := document.New()
	
	docName := projectName(sourceCodeDir) + ".docx"
	err := recursionDir2(doc, sourceCodeDir, docName, exceptDirs, exts)
	
	return err
}

// 主函数
func main() {
	//rootDir := "/home/windness/workspace/go/src/bookingSystem"
	//recursionDir(rootDir)
	
	//fileContent, _ := fileContent("/home/windness/workspace/go/src/bookingSystem/main.go")
	//fmt.Println(fileContent)
	
	//fileExt, _ := fileExt("/home/windness/workspace/go/src/bookingSystem/main.go")
	//fmt.Println(fileExt)

	//doc := document.New()
	//createDocWithOnelineTable(doc, "title", "content1", "filename.docx")
	//createDocWithOnelineTable(doc, "title", "content2", "filename.docx")
	exts := []string{".go", ".html", ".json", ".conf", ".tpl"}
	//fmt.Println(isNeededFile("/home/windness/workspace/go/src/bookingSystem/main.go", exts))
	//sourceCodeDir := "/home/windness/workspace/go/src/bookingSystem/"
	sourceCodeDir := "/home/windness/workspace/go/src/github.com/gwuhaolin/livego/"
	exceptDirs := []string{"static", "views"}
	//sourceCodeDir := "H:/workspace/golang/booking/"
	//sDir, sBase := path.Split(sourceCodeDir)
	//fmt.Println(sDir, " ", sBase)
	//fmt.Println(projectName("/home/windness/workspace/go/src/bookingSystem/"))
	if err := generateCodeFile(sourceCodeDir, exceptDirs, exts); err != nil {
		fmt.Println("出错了...")
	}

	fmt.Println("生成完成")
}
