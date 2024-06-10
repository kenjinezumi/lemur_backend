package ppt

import (
	"bytes"
	"strconv"

	"github.com/unidoc/unioffice/presentation"
)

type SlideContent struct {
	Insights        []string `json:"insights"`
	Recommendations []string `json:"recommendations"`
	Drivers         []string `json:"drivers"`
	Codes           []string `json:"codes"`
}

func CreatePowerPoint(content map[int]SlideContent) ([]byte, error) {
	ppt, err := presentation.New()
	if err != nil {
		return nil, err
	}

	for slideNo, slideContent := range content {
		slide := ppt.AddSlide()
		addContentToSlide(slide, slideNo, slideContent)
	}

	buf := new(bytes.Buffer)
	if err := ppt.Save(buf); err != nil {
		return nil, err
	}

	return buf.Bytes(), nil
}

func addContentToSlide(slide *presentation.Slide, slideNo int, content SlideContent) {
	// Add title
	titleShape := slide.AddTitle()
	titleShape.SetText("Slide " + strconv.Itoa(slideNo))

	// Add insights
	if len(content.Insights) > 0 {
		p := slide.AddParagraph()
		p.SetText("Insights:")
		for _, insight := range content.Insights {
			p.AddRun().SetText("- " + insight)
		}
	}

	// Add recommendations
	if len(content.Recommendations) > 0 {
		p := slide.AddParagraph()
		p.SetText("Recommendations:")
		for _, recommendation := range content.Recommendations {
			p.AddRun().SetText("- " + recommendation)
		}
	}

	// Add drivers
	if len(content.Drivers) > 0 {
		p := slide.AddParagraph()
		p.SetText("Drivers:")
		for _, driver := range content.Drivers {
			p.AddRun().SetText("- " + driver)
		}
	}

	// Add codes
	if len(content.Codes) > 0 {
		p := slide.AddParagraph()
		p.SetText("Codes:")
		for _, code := range content.Codes {
			p.AddRun().SetText("- " + code)
		}
	}
}
