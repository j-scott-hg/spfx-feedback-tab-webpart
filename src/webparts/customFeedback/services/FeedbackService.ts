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

  private async getRequestDigest(): Promise<string> {
    const response: SPHttpClientResponse = await this.context.spHttpClient.post(
      `${this.siteUrl}/_api/contextinfo`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-Type': 'application/json;odata=nometadata'
        }
      }
    );
    const data = await response.json();
    return data.FormDigestValue;
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
          headers: { 'Accept': 'application/json;odata=nometadata' }
        }
      );
      return response.ok;
    } catch {
      return false;
    }
  }

  private async createList(): Promise<void> {
    const digest = await this.getRequestDigest();

    const createResponse: SPHttpClientResponse = await this.context.spHttpClient.post(
      `${this.siteUrl}/_api/web/lists`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-Type': 'application/json;odata=nometadata',
          'X-RequestDigest': digest
        },
        body: JSON.stringify({
          Title: LIST_NAME,
          BaseTemplate: 100,
          Description: 'Stores feedback submitted by users via the Feedback web part.'
        })
      }
    );

    if (!createResponse.ok) {
      throw new Error(`Failed to create list: ${createResponse.statusText}`);
    }

    await this.addListColumns(digest);
  }

  private async addListColumns(digest: string): Promise<void> {
    const baseUrl = `${this.siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(LIST_NAME)}')/fields`;
    const headers = {
      'Accept': 'application/json;odata=nometadata',
      'Content-Type': 'application/json;odata=nometadata',
      'X-RequestDigest': digest
    };

    const columns = [
      { FieldTypeKind: 3, Title: 'CommentText', Required: false },
      { FieldTypeKind: 9, Title: 'Rating', Required: false },
      { FieldTypeKind: 2, Title: 'PageURL', Required: false },
      { FieldTypeKind: 2, Title: 'PageTitle', Required: false },
      { FieldTypeKind: 2, Title: 'SiteURL', Required: false },
      { FieldTypeKind: 4, Title: 'SubmittedDateTime', Required: false }
    ];

    for (const col of columns) {
      await this.context.spHttpClient.post(baseUrl, SPHttpClient.configurations.v1, {
        headers,
        body: JSON.stringify(col)
      });
    }
  }

  public async submitFeedback(submission: IFeedbackSubmission): Promise<void> {
    const digest = await this.getRequestDigest();

    const response: SPHttpClientResponse = await this.context.spHttpClient.post(
      `${this.siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(LIST_NAME)}')/items`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-Type': 'application/json;odata=nometadata',
          'X-RequestDigest': digest
        },
        body: JSON.stringify(submission)
      }
    );

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Failed to submit feedback: ${errorText}`);
    }
  }
}
