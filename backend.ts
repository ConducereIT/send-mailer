import {GenezioDeploy, GenezioMethod} from "@genezio/types";

@GenezioDeploy()
export class BackendService {
  constructor() {
  }

  @GenezioMethod()
  async getHello(): Promise<string> {
    return 'Hello World';
  }
}
