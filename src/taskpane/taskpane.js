/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

// eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiJ9.eyJhdWQiOiIxIiwianRpIjoiMDhlMWJlYWYwOTMwNmY0N2E5NTUzNTZlOTM4NDQ1NGVhM2QxZTM3NjUxNWQ2NmI0YjU4OTczNjRlZDA5MjAzMjc3ODRjMDE5ZDQyZWRhZDciLCJpYXQiOjE3NTAyNTY1NTIuODgyNDc5LCJuYmYiOjE3NTAyNTY1NTIuODgyNDgxLCJleHAiOjQ5MDU5MzAxNTIuODc3ODcsInN1YiI6IjcyMjM1Mzc0Iiwic2NvcGVzIjpbInRhc2sucmVhZCIsInRhc2sud3JpdGUiXX0.CS-UzOI0Kuryfm74s0rEGagYfQrWw3PyAmo3fMI3As0mCDasSJH5cdg7MNvM2AMn2X3T9EBolvX8kTm9bV8ab7hcIwvxXqxEyVVsYznaqaJJO2AhvR2U843cOB1Dy8UcSc8fC6Jp7hQSSeWduQPdTWvE9TivWyY5LjFbU-wlKwMeHy9eEjkG5tCERM-stYwGxypKQycnkl208w6PQKUQCI6ZSeHas6egOb62Gemhqy_EXelTj0dvCkqe6LoEJrXrV9aDTXSVWvKef5MDsFYpNAofx1iioGgqekiVJ3apifr-Uenmt-QxwgdYV-8yU5wNa-jFt4D9wmbX9NsE_IJjJYLLd6Uwpk1EVADPkA-fzB8G2-yKYZ00AK078tEnZLCDWCSUpPSK9uD6JerlSfiOZxWfizSBQ4JSmUDa_JUMtzKmablg4zmkBo6ZKH7QWtX9_ItHt4ihrR_8gtZhQn5J37kxgc_Di_zYQZaxdm0slKoI-4YbOyf_HxYfve3KpxhxV2g0NqVeJtMd5KwVI2L9NBhI8o0bgUoDQuG_VUK8W5sQyizhV_TT-txUxBWhlnhUjJEIFpeo-9tDybw2xlY1S7BQm5FH89EAVt7miGLO0wfNTx5V81iwP2H3pIHvLObEL1l1oQ2dQz4jRQTBVfV2HMNuj6XetzwbMdhTE1_LKPs
// eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiJ9.eyJhdWQiOiIxIiwianRpIjoiMjAyMjQxYjIzMmU0YmI2NWFiNWUxMDRiMTU3MjhiYTlmNjU4MGJjN2E5ZDlhNDYxN2IwNzUyOGIyY2FjOTNhM2JjNGZlZDdjMWY4ODAwZjkiLCJpYXQiOjE3NTAzNTExNTYuODAxMTE2LCJuYmYiOjE3NTAzNTExNTYuODAxMTE3LCJleHAiOjQ5MDYwMjQ3NTYuNzk1MDksInN1YiI6IjcyMjM1Mzc0Iiwic2NvcGVzIjpbInRhc2sucmVhZCIsInRhc2sud3JpdGUiXX0.pnkw9UnWkzO5Hz22UozOXtx1TTRRtXkRCggpVE2dS7DFo6Rly5lFBQdVOizkbmEEs9hZuJQz6U5UN23sJ8rDd5UFosYaMcpO7-bejHhwRFizME9i2EMsWzT9JsfNIi2WKtKnspjVExmIu7jHzIyWYqUijjYv4JZ2AKtE-AE15wm0KVl8mafz8QyYemWXFXLrdhDNsjng1CoqRZ2UXtGsNZFOdj522wBv5XQi0aLoct4pSzP7qdCqhkQMbNw5rorme_q6Tja90C8JjhMAUiNhDGObsdSWhefCVVxB0o1uoZMS4V0jGpH-_VKKeg2lgd1sLSqQXwi1ErzZRp5QHnRlUhRlNFuz0d1PIdHTMlCqM7J4FevslORZVdKY7I179OZkFXS3o51CHFDW6Ymvg-lEhrYIdSIk0KyRySjoykYa8ITex5Tnoy3qwP-SaGzYcydiaRddGJBeWxKSiPgfO1_r83DdI4jx7pROPj10Vf5Byl3K89mQ42EUBBl7f0oPfRCzo6_qddz8uvIQi5zm8oLyMoejLgFyJaLfOSTArr6GmHhyVn05oHL9oQJPLEXs8mVyt5xj9NkrIFsT5pRbqR0D03qInTXGJOpQkBud3g8hWmUEv0AgT64lunLT7uBZafiuXb-421fbveGaBNAdull5wW5dOIzqSvL27r-vPu8CMeQ

// The initialize function must be run each time a new page is loaded
// Office.onReady(() => {
//   document.getElementById("sideload-msg").style.display = "none";
//   document.getElementById("app-body").style.display = "flex";
//   document.getElementById("run").onclick = run;
// });

// export async function run() {
//   try {
//     await Excel.run(async (context) => {
//       /**
//        * Insert your Excel code here
//        */
//       const range = context.workbook.getSelectedRange();

//       // Read the range address
//       range.load("address");

//       // Update the fill color
//       range.format.fill.color = "yellow";

//       await context.sync();
//       console.log(`The range address was ${range.address}.`);
//     });
//   } catch (error) {
//     console.error(error);
//   }
// }





Office.onReady(() => {
  document.getElementById("convertBtn").onclick = async () => {
    const fileInput = document.getElementById("uploadDocx");
    const status = document.getElementById("status");
    const file = fileInput.files[0];

    if (!file) {
      status.innerText = "❌ Please select a .docx file first.";
      return;
    }

    try {
      status.innerText = "🔄 Creating conversion job...";

      const apiKey = "eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiJ9.eyJhdWQiOiIxIiwianRpIjoiZGJmZTliNWJiOTBlMTY5OGNlZTdmZmI3NzIyZTg1MjgwMzQyODU3ZTYxY2M1MmZlNDAzMGVmZjdmZGM3YjdmZmRhZmE3NjYyMGQ0ODQyNGIiLCJpYXQiOjE3NTA1OTk2MjYuMjI3MTMyLCJuYmYiOjE3NTA1OTk2MjYuMjI3MTM0LCJleHAiOjQ5MDYyNzMyMjYuMjIzNTEzLCJzdWIiOiI3MjIzNTM3NCIsInNjb3BlcyI6WyJ0YXNrLnJlYWQiLCJ0YXNrLndyaXRlIl19.hmJsJlC28R-XMn8PvV10ZTnZD425ZUh4cBTpIsXMpFJ4FijCqgoibTKofRr7XgXHt8lkeMj9a3IJf8tfuLxx_U8OavhlxMRAG7MPBefKoaJynmp_ZxgI1GVS0W7NLp2jyo3iFnxSQa-H9qkWL8_TTfot72uGQsurnHy3LcwC-hLZfE63YMnndjSn3jFCGSCNOSELIIWosRZNkoHTegdRWL6rcKoIYcfi4Ff7EfwLL2w7-oTnSzICu-Lg-XnmoXRp_jf2AYEpEkvXnlOAKM86pBp3h3DlAcZZOsYr_WIJ8e8Ti4hmlAgwqD9yN76DZeQ65Wq9kBF2LbKiV8fh2WLERHDlA-0OKN92mo4qNvCsfThYgrRXcnNY9oMnYPHewaPwwT6PHv7DJGdLTHZUrtWc6mLLq1XpBzIMEzv1_qbrguot01lwrkkVACDZvnXaUPJ3plM04j6XlKf6qFPyoDNxR-ThbDjxpXqnCM57RDjE1E5ow3cXW2a_5jMFhCjvpMTyUkapzmXPDKF9rYTog4hw5aA6PRzIyS0yMIZB4VaxGM4S-tPwCJNdSdFSz7bssdTVnSf_h0GvlgUcp3en6qT6KfIIMGP7KAGsSZkB339Fu-7U7pJj7wosd1myL1-FXhauZI_YmL5W34SD6tLzvLbEfYfZ0pm9-lsSVXpMqgHlHLM"; // Replace this

      // Step 1: Create conversion job
      const jobRes = await fetch("https://api.cloudconvert.com/v2/jobs", {
        method: "POST",
        headers: {
          "Authorization": `Bearer ${apiKey}`,
          "Content-Type": "application/json"
        },
        body: JSON.stringify({
          tasks: {
            upload: { operation: "import/upload" },
            convert: {
              operation: "convert",
              input: "upload",
              input_format: "docx",
              output_format: "pdf"
            },
            export: { operation: "export/url", input: "convert" }
          }
        })
      });

      const job = await jobRes.json();
      const uploadTask = Object.values(job.data.tasks).find(t => t.name === "upload");

      // Step 2: Upload file
      status.innerText = "📤 Uploading DOCX file...";
      const formData = new FormData();
      for (const key in uploadTask.result.form.parameters) {
        formData.append(key, uploadTask.result.form.parameters[key]);
      }
      formData.append("file", file);

      await fetch(uploadTask.result.form.url, {
        method: "POST",
        body: formData
      });

      // Step 3: Wait for job completion
      status.innerText = "⏳ Waiting for conversion...";
      let exportTask;
      let done = false;
      while (!done) {
        const poll = await fetch(`https://api.cloudconvert.com/v2/jobs/${job.data.id}`, {
          headers: { Authorization: `Bearer ${apiKey}` }
        });
        const updatedJob = await poll.json();
        done = updatedJob.data.status === "finished";
        exportTask = updatedJob.data.tasks.find(t => t.name === "export");
        if (!done) await new Promise(r => setTimeout(r, 3000));
      }

      // Step 4: Show download link
      const fileUrl = exportTask.result.files[0].url;
      status.innerHTML = `✅ Done! <a href="${fileUrl}" target="_blank">Download PDF</a>`;
    } catch (err) {
      console.error(err);
      status.innerText = "❌ Conversion failed. Check the console for errors.";
    }
  };
});





// function logToDebugPanel(message) {
//   const debug = document.getElementById("debug");
//   if (debug) {
//     debug.textContent += `${message}\n`;
//   }
// }

// Office.onReady(() => {
//   document.getElementById("convertBtn").onclick = async () => {
//     const fileInput = document.getElementById("uploadDocx");
//     const status = document.getElementById("status");
//     const file = fileInput.files[0];

//     if (!file) {
//       status.innerText = "❌ Please select a .docx file first.";
//       logToDebugPanel("No file selected.");
//       return;
//     }

//     try {
//       status.innerText = "🔄 Creating conversion job...";
//       logToDebugPanel("Creating CloudConvert job...");

//       const apiKey = "eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiJ9.eyJhdWQiOiIxIiwianRpIjoiZGJmZTliNWJiOTBlMTY5OGNlZTdmZmI3NzIyZTg1MjgwMzQyODU3ZTYxY2M1MmZlNDAzMGVmZjdmZGM3YjdmZmRhZmE3NjYyMGQ0ODQyNGIiLCJpYXQiOjE3NTA1OTk2MjYuMjI3MTMyLCJuYmYiOjE3NTA1OTk2MjYuMjI3MTM0LCJleHAiOjQ5MDYyNzMyMjYuMjIzNTEzLCJzdWIiOiI3MjIzNTM3NCIsInNjb3BlcyI6WyJ0YXNrLnJlYWQiLCJ0YXNrLndyaXRlIl19.hmJsJlC28R-XMn8PvV10ZTnZD425ZUh4cBTpIsXMpFJ4FijCqgoibTKofRr7XgXHt8lkeMj9a3IJf8tfuLxx_U8OavhlxMRAG7MPBefKoaJynmp_ZxgI1GVS0W7NLp2jyo3iFnxSQa-H9qkWL8_TTfot72uGQsurnHy3LcwC-hLZfE63YMnndjSn3jFCGSCNOSELIIWosRZNkoHTegdRWL6rcKoIYcfi4Ff7EfwLL2w7-oTnSzICu-Lg-XnmoXRp_jf2AYEpEkvXnlOAKM86pBp3h3DlAcZZOsYr_WIJ8e8Ti4hmlAgwqD9yN76DZeQ65Wq9kBF2LbKiV8fh2WLERHDlA-0OKN92mo4qNvCsfThYgrRXcnNY9oMnYPHewaPwwT6PHv7DJGdLTHZUrtWc6mLLq1XpBzIMEzv1_qbrguot01lwrkkVACDZvnXaUPJ3plM04j6XlKf6qFPyoDNxR-ThbDjxpXqnCM57RDjE1E5ow3cXW2a_5jMFhCjvpMTyUkapzmXPDKF9rYTog4hw5aA6PRzIyS0yMIZB4VaxGM4S-tPwCJNdSdFSz7bssdTVnSf_h0GvlgUcp3en6qT6KfIIMGP7KAGsSZkB339Fu-7U7pJj7wosd1myL1-FXhauZI_YmL5W34SD6tLzvLbEfYfZ0pm9-lsSVXpMqgHlHLM"; // replace with your actual key

//       // Step 1: Create job
//       const jobRes = await fetch("https://api.cloudconvert.com/v2/jobs", {
//         method: "POST",
//         headers: {
//           "Authorization": `Bearer ${apiKey}`,
//           "Content-Type": "application/json"
//         },
//         body: JSON.stringify({
//           tasks: {
//             upload: { operation: "import/upload" },
//             convert: {
//               operation: "convert",
//               input: "upload",
//               input_format: "docx",
//               output_format: "pdf"
//             },
//             export: { operation: "export/url", input: "convert" }
//           }
//         })
//       });

//       const job = await jobRes.json();
//       logToDebugPanel("Job created: " + JSON.stringify(job, null, 2));

//       const uploadTask = Object.values(job.data.tasks).find(t => t.name === "upload");

//       // Step 2: Upload file
//       status.innerText = "📤 Uploading DOCX file...";
//       logToDebugPanel("Uploading file...");

//       const formData = new FormData();
//       for (const key in uploadTask.result.form.parameters) {
//         formData.append(key, uploadTask.result.form.parameters[key]);
//       }
//       formData.append("file", file);

//       await fetch(uploadTask.result.form.url, {
//         method: "POST",
//         body: formData
//       });

//       // Step 3: Poll job status
//       status.innerText = "⏳ Waiting for conversion...";
//       logToDebugPanel("Waiting for conversion...");

//       let exportTask;
//       let done = false;
//       while (!done) {
//         const poll = await fetch(`https://api.cloudconvert.com/v2/jobs/${job.data.id}`, {
//           headers: { Authorization: `Bearer ${apiKey}` }
//         });
//         const updatedJob = await poll.json();
//         done = updatedJob.data.status === "finished";
//         exportTask = updatedJob.data.tasks.find(t => t.name === "export");

//         logToDebugPanel(`Polling job status: ${updatedJob.data.status}`);
//         if (!done) await new Promise(r => setTimeout(r, 3000));
//       }

//       // Step 4: Provide download link
//       const fileUrl = exportTask.result.files[0].url;
//       status.innerHTML = `✅ Done! <a href="${fileUrl}" target="_blank">Download PDF</a>`;
//       logToDebugPanel("Conversion completed: " + fileUrl);

//     } catch (err) {
//       console.error(err);
//       status.innerText = "❌ Conversion failed. Check the console for errors.";
//       logToDebugPanel("Error: " + err.message);
//     }
//   };
// });

