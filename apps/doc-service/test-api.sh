#!/bin/bash

echo "Testing stream-docx..."
curl -s 'http://127.0.0.1:3001/api/v1/docs/stream-docx?docId=mock.docx' | wc -c

echo "Testing better-stream-docx..."
timeout 5 curl -s 'http://127.0.0.1:3001/api/v1/docs/better-stream-docx?docId=mock.docx' | wc -c

echo "Done"
