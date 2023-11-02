import { INode, INodeData, INodeParams } from "../../../src/Interface";
import { TextSplitter } from "langchain/text_splitter";
import { DocxLoader } from "langchain/document_loaders/fs/docx";
import { fetchFileFromUrl } from "../../../src/utils";

class Docx_DocumentLoaders implements INode {
  label: string;
  name: string;
  version: number;
  description: string;
  type: string;
  icon: string;
  category: string;
  baseClasses: string[];
  inputs: INodeParams[];

  constructor() {
    this.label = "Docx File";
    this.name = "docxFile";
    this.version = 1.0;
    this.type = "Document";
    this.icon = "Docx.png";
    this.category = "Document Loaders";
    this.description = `Load data from DOCX files`;
    this.baseClasses = [this.type];
    this.inputs = [
      {
        label: "Docx File",
        name: "docxFile",
        type: "file",
        fileType: ".docx",
      },
      {
        label: "Text Splitter",
        name: "textSplitter",
        type: "TextSplitter",
        optional: true,
      },
      {
        label: "Metadata",
        name: "metadata",
        type: "json",
        optional: true,
        additionalParams: true,
      },
    ];
  }

  async init(nodeData: INodeData): Promise<any> {
    const textSplitter = nodeData.inputs?.textSplitter as TextSplitter;
    const docxFileUrls = nodeData.inputs?.docxFile as string;
    const metadata = nodeData.inputs?.metadata;
    const mimeType =
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

    let alldocs = [];
    let files: string[] = [];

    if (docxFileUrls.startsWith("[") && docxFileUrls.endsWith("]")) {
      const filesArray: string[] = JSON.parse(docxFileUrls);
      filesArray.forEach(async (url) => {
        const base64 = await fetchFileFromUrl(url, mimeType);
        files.push(base64);
      });
    } else {
      const base64 = await fetchFileFromUrl(docxFileUrls, mimeType);
      files = [base64];
    }

    for (const file of files) {
      const splitDataURI = file.split(",");
      splitDataURI.pop();
      const bf = Buffer.from(splitDataURI.pop() || "", "base64");
      const blob = new Blob([bf]);
      const loader = new DocxLoader(blob);

      if (textSplitter) {
        const docs = await loader.loadAndSplit(textSplitter);
        alldocs.push(...docs);
      } else {
        const docs = await loader.load();
        alldocs.push(...docs);
      }
    }

    if (metadata) {
      const parsedMetadata =
        typeof metadata === "object" ? metadata : JSON.parse(metadata);
      let finaldocs = [];
      for (const doc of alldocs) {
        const newdoc = {
          ...doc,
          metadata: {
            ...doc.metadata,
            ...parsedMetadata,
          },
        };
        finaldocs.push(newdoc);
      }
      return finaldocs;
    }

    return alldocs;
  }
}

module.exports = { nodeClass: Docx_DocumentLoaders };
