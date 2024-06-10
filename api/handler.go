package api

import (
	"bytes"
	"encoding/json"
	"io/ioutil"
	"lemur_backend/drive"
	"lemur_backend/ppt"
	"log"
	"net/http"

	"github.com/gin-gonic/gin"
)

type RequestPayload struct {
	QuarterNo string `json:"quarter_no"`
	YearNo    string `json:"year_no"`
	FileID    string `json:"file_id"`
}

type APIResponse map[int]ppt.SlideContent

func HandleGenerate(c *gin.Context) {
	var payload RequestPayload
	if err := c.BindJSON(&payload); err != nil {
		c.JSON(http.StatusBadRequest, gin.H{"error": err.Error()})
		return
	}

	go processRequest(payload)

	c.JSON(http.StatusAccepted, gin.H{"status": "Processing"})
}

func processRequest(payload RequestPayload) {
	apiResponse, err := callAPI(payload.QuarterNo, payload.YearNo)
	if err != nil {
		log.Println("Error calling API:", err)
		return
	}

	pptx, err := ppt.CreatePowerPoint(apiResponse)
	if err != nil {
		log.Println("Error creating PowerPoint:", err)
		return
	}

	fileID, err := drive.UploadFileToDrive(payload.FileID, pptx)
	if err != nil {
		log.Println("Error uploading file to Google Drive:", err)
		return
	}

	log.Println("File successfully uploaded with ID:", fileID)
}

func callAPI(quarterNo, yearNo string) (APIResponse, error) {
	url := "http://34.90.192.243/deman_gen_insights"
	payload := map[string]string{
		"quarter_no": quarterNo,
		"year_no":    yearNo,
	}
	jsonData, _ := json.Marshal(payload)

	resp, err := http.Post(url, "application/json", bytes.NewBuffer(jsonData))
	if err != nil {
		return nil, err
	}
	defer resp.Body.Close()

	body, _ := ioutil.ReadAll(resp.Body)

	var apiResponse APIResponse
	if err := json.Unmarshal(body, &apiResponse); err != nil {
		return nil, err
	}

	return apiResponse, nil
}
