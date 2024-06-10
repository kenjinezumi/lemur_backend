package main

import (
	"lemur_backend/api"

	"github.com/gin-gonic/gin"
)

func main() {
	r := gin.Default()
	r.POST("/generate", api.HandleGenerate)
	r.Run() // listen and serve on 0.0.0.0:8080
}
