![banner](./images/banner.png)

## Convert ASR JSON To Word Document

#### Overview

This Python3 application will process the results of a synchronous **Amazon Transcribe** job, in either Transcribe standard mode or **Transcribe Call Analytics** (TCA) mode, or **Amazon Bedrock Data Automation Audio** (BDA) job and will turn it into a Microsoft Word document that contains a turn-by-turn transcript from each speaker.  It can handle processing a local JSON output file, or it can dynamically query the Amazon Transcribe service to download the JSON.  It works in one of two different modes:

- **Transcribe Standard Mode** - this can optionally call Amazon Comprehend to generate sentiment for each turn of the conversation.  This mode can handle either speaker-separated or channel-separated audio files
- **Transcribe Call Analytics Mode** - using the Call Analytics feature of Amazon Transcribe, the Word document will present all of the analytical data in either a tabular or graphical form
- **BDA Audio Mode** - using the standard BDA Audio project, and optionally the contact center custom blueprint.  The standard project includes general transcript-related features, but the custom blueprint includes additional call-level features



#### Features

The following table summarise which features are available in each mode.  It should be noted that some missing items in BDA Audio can be done with a non-standard custom blueprint.  Also, not all features are supported by this demo, such as *Dominant Language Detection*.

| **Feature**                           | **Transcribe**   | Call Analytics | BDA Audio        |
| ------------------------------------- | ---------------- | -------------- | ---------------- |
| ***Standard Call Characteristics***   |                  |                |                  |
| Job information [1]                   | ✅                | ✅              | ❌                |
| Word-level confidence scores          | ✅                | ✅              | ❌                |
| Word-level timings                    | ✅                | ✅              | ✅                |
| ***Sentiment Analysis***              |                  |                |                  |
| Call-level sentiment [2]              | ❌                | ✅              | ✅                |
| Speaker sentiment trend [2] [3]       | *Via Comprehend* | ✅              | ✅                |
| Turn-level sentiment                  | *Via Comprehend* | ✅              | *Via Comprehend* |
| Turn-level sentiment scores           | *Via Comprehend* | ❌              | *Via Comprehend* |
| ***Call Characteristics***            |                  |                |                  |
| Call issue detection [2] [4]          | ❌                | ✅              | ✅                |
| Call non-talk ("silent") time         | ❌                | ✅              | ❌                |
| Category detection [5]                | ❌                | ✅              | ✅                |
| Entity detection [2]                  | ❌                | ❌              | ✅                |
| Generative call summarisation [4] [6] | ❌                | ✅              | ✅                |
| Named speaker identification          | ❌                | ✅              | ❌                |
| Speaker interruptions                 | ❌                | ✅              | ❌                |
| Speaker talk time                     | ❌                | ✅              | ❌                |
| Speaker volume                        | ❌                | ✅              | ❌                |
| Toxicity detection                    | ✅                | ❌              | ✅                |

*[1] The completed job, if it ever existed, can be manually deleted, and is not guaranteed to persist over time*

*[2] This feature in BDA Audio is only possibly via a custom blueprint*

*[3] Speaker sentiment trend in Transcribe Call Analytics only provides data points for each quarter of the call per speaker*

*[4] This feature in Transcribe or Transcribe Call Analytics is only available on synchronous streams*

*[5] Categories must first be defined by the customer, and only those defined when the job is executed will be reported*

*[6] This feature in Transcribe or Transcribe Call Analytics is not available in all supported languages*



#### Usage

##### Prerequisites

This application relies upon three external python libraries, which you will need to install onto the system that you wise to deploy this application to.  They are as follows:

- python-docx
- scipy
- matplotlib

These should all be installed using the relevant tool for your target platform - typically this would be via `pip`, the Python package manager, but could be via `yum`, `apt-get` or something else.  Please consult your platform's Python documentation for more information.

Additionally, as the Python code will call APIs in Amazon Transcribe and, optionally, Amazon Comprehend, the target platform will need to have access to AWS access keys or an IAM role that gives access to the following API calls:

- Amazon Transcribe - *GetTranscriptionJob()* and *GetCallAnalyticsJob()*

- Amazon Comprehend - *DetectSentiment()*

  

#### Sample files

The repository contains the following sample files in the `sample-data` folder:

- **example-call.wav** - an example two-channel call audio file
- **example-call.json** - the result from Amazon Transcribe when the example audio file is processed in Call Analytics mode
- **example-call.docx** - the output document generated by this application against a completed Amazon Transcribe Call Analytics job using the example audio file.  The next section describes this file structure in more detail
- **example-call-redacted.wav** - the example call with all PII redacted, which can be output by Call Analytics if you enable PII and request that results are delivered to your own S3 bucket



#### Mode-specific output

<kbd>[**Transcribe & Call Analytics**](./README-transcribe.md)</kbd>      <kbd>[**Bedrock Data Automation Audio**](./README-bda.md)</kbd>

## Security	

See [CONTRIBUTING](CONTRIBUTING.md#security-issue-notifications) for more information.

## License

This library is licensed under the MIT-0 License. See the LICENSE file.

