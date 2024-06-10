package drive

import (
	"bytes"
	"context"

	"google.golang.org/api/drive/v3"
	"google.golang.org/api/option"
)

func UploadFileToDrive(folderID string, pptx []byte) (string, error) {
	ctx := context.Background()
	srv, err := drive.NewService(ctx, option.WithCredentialsFile("/app/service-account-key.json"))
	if err != nil {
		return "", err
	}

	file := &drive.File{
		Name:    "GeneratedReport.pptx",
		Parents: []string{folderID},
	}

	pptxReader := bytes.NewReader(pptx)
	res, err := srv.Files.Create(file).Media(pptxReader).Do()
	if err != nil {
		return "", err
	}

	return res.Id, nil
}
