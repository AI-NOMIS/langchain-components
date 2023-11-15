import {
  ICommonObject,
  INode,
  INodeData,
  INodeOptionsValue,
  INodeParams,
} from "../../../src/Interface";
import { getBaseClasses } from "../../../src/utils";
import { DynamicStructuredTool } from "./core";
import { z } from "zod";
// @ts-ignore
import { dbClient } from "../../../../../../src/shared/client";
// Note: the above import refers to the nocodingai_backend repo's db instance. There is
//       an additional ../ in the from string because this file will be compiled into
//       the 'dist' folder that's one level deeper than this current file. Need to find
//       better alternative in the future. This is a temp solution.

class CustomTool_Tools implements INode {
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
    this.label = "Custom Tool";
    this.name = "customTool";
    this.version = 1.0;
    this.type = "CustomTool";
    this.icon = "customtool.svg";
    this.category = "Tools";
    this.description = `Use custom tool you've created in Flowise within chatflow`;
    this.inputs = [
      {
        label: "Select Tool",
        name: "selectedTool",
        type: "asyncOptions",
        loadMethod: "listTools",
      },
    ];
    this.baseClasses = [
      this.type,
      "Tool",
      ...getBaseClasses(DynamicStructuredTool),
    ];
  }

  //@ts-ignore
  loadMethods = {
    async listTools(
      _: INodeData,
      options: ICommonObject
    ): Promise<INodeOptionsValue[]> {
      const returnData: INodeOptionsValue[] = [];

      const tools = await dbClient.chatTool.findMany();

      for (let i = 0; i < tools.length; i += 1) {
        const data = {
          label: tools[i].name,
          name: tools[i].id,
          description: tools[i].description,
        } as INodeOptionsValue;
        returnData.push(data);
      }
      return returnData;
    },
  };

  async init(
    nodeData: INodeData,
    _: string,
    options: ICommonObject
  ): Promise<any> {
    const selectedToolId = nodeData.inputs?.selectedTool as string;
    const customToolFunc = nodeData.inputs?.customToolFunc as string;

    try {
      const tool = await dbClient.chatTool.findUnique({
        where: {
          id: selectedToolId,
        },
      });

      if (tool == null) throw new Error(`Tool ${selectedToolId} not found`);
      if (tool.schema == null)
        throw new Error(`Tool schema for tool id ${selectedToolId} not found`);
      if (tool.func == null)
        throw new Error(
          `Tool function for toold id ${selectedToolId} not found`
        );
      const obj = {
        name: tool.name,
        description: tool.description,
        schema: z.object(convertSchemaToZod(tool.schema)),
        code: tool.func,
      };
      if (customToolFunc) obj.code = customToolFunc;
      return new DynamicStructuredTool(obj);
    } catch (e) {
      throw new Error(e);
    }
  }
}

const convertSchemaToZod = (schema: string) => {
  try {
    const parsedSchema = JSON.parse(schema);
    const zodObj: any = {};
    for (const sch of parsedSchema) {
      if (sch.type === "string") {
        if (sch.required)
          z.string({ required_error: `${sch.property} required` }).describe(
            sch.description
          );
        zodObj[sch.property] = z.string().describe(sch.description);
      } else if (sch.type === "number") {
        if (sch.required)
          z.number({ required_error: `${sch.property} required` }).describe(
            sch.description
          );
        zodObj[sch.property] = z.number().describe(sch.description);
      } else if (sch.type === "boolean") {
        if (sch.required)
          z.boolean({ required_error: `${sch.property} required` }).describe(
            sch.description
          );
        zodObj[sch.property] = z.boolean().describe(sch.description);
      }
    }
    return zodObj;
  } catch (e) {
    throw new Error(e);
  }
};

module.exports = { nodeClass: CustomTool_Tools };
