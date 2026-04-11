import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IFeedbackSubmission } from '../components/IFeedbackTypes';

const LIST_NAME = 'Feedback Responses';

export class FeedbackService {
  private context: WebPartContext;
  private siteUrl: string;

  constructor(context: WebPartContext, siteUrl: string) {
    this.context = context;
    this.siteUrl = siteUrl;
  }

  // SPHttpClient automatically handles the request digest (X-RequestDigest) for
  // state-changing calls — no manual contextinfo fetch needed in SPFx.
  private getPostOptions(body: object): { headers: Record<string, string>; body: string } {
    return {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-Type': 'application/json;odata=nometadata',
        'odata-version': ''
      },
      body: JSON.stringify(body)
    };
  }

  public async ensureList(): Promise<void> {
    const exists = await this.listExists();
    if (!exists) {
      await this.createList();
    }
  }

  private async listExists(): Promise<boolean> {
    try {
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(LIST_NAME)}')`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'odata-version': ''
          }
        }
      );
      return response.ok;
    } catch {
      return false;
    }
  }

  private async createList(): Promise<void> {
    const createResponse: SPHttpClientResponse = await this.context.spHttpClient.post(
      `${this.siteUrl}/_api/web/lists`,
      SPHttpClient.configurations.v1,
      this.getPostOptions({
        Title: LIST_NAME,
        BaseTemplate: 100,
        Description: 'Stores feedback submitted by users via the Feedback Tab web part.'
      })
    );

    if (!createResponse.ok) {
      throw new Error(`Failed to create list: ${createResponse.statusText}`);
    }

    await this.addListColumns();
  }

  private async addListColumns(): Promise<void> {
    const baseUrl = `${this.siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(LIST_NAME)}')/fields`;

    const columns = [
      { FieldTypeKind: 3, Title: 'CommentText', Required: false },
      { FieldTypeKind: 9, Title: 'Rating', Required: false },
      { FieldTypeKind: 2, Title: 'PageURL', Required: false },
      { FieldTypeKind: 2, Title: 'PageTitle', Required: false },
      { FieldTypeKind: 2, Title: 'SiteURL', Required: false },
      { FieldTypeKind: 4, Title: 'SubmittedDateTime', Required: false }
    ];

    for (const col of columns) {
      await this.context.spHttpClient.post(
        baseUrl,
        SPHttpClient.configurations.v1,
        this.getPostOptions(col)
      );
    }
  }

  public async submitFeedback(submission: IFeedbackSubmission): Promise<void> {
    const response: SPHttpClientResponse = await this.context.spHttpClient.post(
      `${this.siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(LIST_NAME)}')/items`,
      SPHttpClient.configurations.v1,
      this.getPostOptions(submission)
    );

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Failed to submit feedback: ${errorText}`);
    }
  }
}
