#(Invoke-WebRequest "https://script-vn.github.io/cmd/note.txt").ParsedHtml.body.innerText
Invoke-RestMethod -Uri "https://script-vn.github.io/cmd/note.txt"
