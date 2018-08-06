function checkStatus(response) {
  if (response.status >= 200 && response.status < 300) {
    return response
  } else {
    var error = new Error(response.statusText)
    error.response = response
    throw error
  }
}

function parseJSON(response) {
  return response.json()
}

document.getElementById('testBtn')
  .addEventListener('click' , function () {
    let contractNo = document.getElementById('contractNo').value
    console.log(contractNo)
    let options = {
      contractNo
    }
    fetch('/word', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(options)
    })
      .then(checkStatus)
      .then(parseJSON)
      .then(res => {
        window.open(window.location.href + res.path)
      })
  })
