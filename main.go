package main

import (
	"fmt"
	"github.com/joho/godotenv"
	"github.com/unidoc/unioffice/common/license"
	"github.com/unidoc/unioffice/document"
	"os"
	"strings"
)

func init() {
	loadKeyError := godotenv.Load()
	if loadKeyError != nil {
		fmt.Println("Please use the correct API_KEY")
		return
	}
	err := license.SetMeteredKey(os.Getenv("UNIOFFICE_API_KEY"))
	if err != nil {
		panic(err)
	}
}

func main() {
	// open Word.docx
	doc, err := document.Open("./example/example.docx")
	if err != nil {
		fmt.Printf("Cannot open your document")
		return
	}

	// 遍历所有段落
	for _, p := range doc.Paragraphs() {
		// 遍历段落中的所有文本运行
		for _, run := range p.Runs() {
			// 判断文本是否为红色
			if colorRGB := run.Properties().GetColor().AsRGBString(); *colorRGB == "ff0000" {
				// 提取单词
				words := strings.Fields(run.Text())
				// 输出满足条件的单词
				for _, word := range words {
					fmt.Println(word)
				}
			}
		}
	}
}
