import * as React from 'react';
import { SPHttpClient } from '@microsoft/sp-http';
import styles from './RestLists.module.scss';
import type { IRestListsProps } from './IRestListsProps';

interface ISPListInfo {
  Id: string;
  Title: string;
  BaseTemplate: number;
  Hidden: boolean;
  ItemCount: number;
}

interface ISPListResponse {
  value: ISPListInfo[];
}

const RestLists: React.FC<IRestListsProps> = ({ context }) => {
  const [lists, setLists] = React.useState<ISPListInfo[]>([]);
  const [loading, setLoading] = React.useState<boolean>(true);
  const [error, setError] = React.useState<string>('');

  React.useEffect(() => {
    let isMounted: boolean = true;

    const loadLists = async (): Promise<void> => {
      try {
        const endpoint = `${context.pageContext.web.absoluteUrl}/_api/web/lists?$select=Title,Id,BaseTemplate,Hidden,ItemCount&$orderby=Title`;
        const response = await context.spHttpClient.get(endpoint, SPHttpClient.configurations.v1);

        if (!response.ok) {
          throw new Error(`REST request failed: ${response.status} ${response.statusText}`);
        }

        const payload = await response.json() as ISPListResponse;

        if (isMounted) {
          setLists(payload.value || []);
        }
      } catch (requestError) {
        if (isMounted) {
          setError(requestError instanceof Error ? requestError.message : String(requestError));
        }
      } finally {
        if (isMounted) {
          setLoading(false);
        }
      }
    };

    loadLists().catch(() => undefined);

    return () => {
      isMounted = false;
    };
  }, [context]);

  if (loading) {
    return <section className={styles.restLists}>Loading lists (REST)...</section>;
  }

  if (error) {
    return <section className={styles.restLists}>Error: {error}</section>;
  }

  return (
    <section className={styles.restLists}>
      <h3>Lists via SharePoint REST</h3>
      <ul>
        {lists.map((list) => (
          <li key={list.Id}>
            {list.Title} (Items: {list.ItemCount}, Hidden: {String(list.Hidden)}, BaseTemplate: {list.BaseTemplate})
          </li>
        ))}
      </ul>
    </section>
  );
};

export default RestLists;
