import * as React from 'react';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import styles from './GraphLists.module.scss';
import type { IGraphListsProps } from './IGraphListsProps';

interface IGraphListInfo {
  id: string;
  displayName: string;
  list?: {
    template?: string;
  };
}

interface IGraphListResponse {
  value: IGraphListInfo[];
}

const GraphLists: React.FC<IGraphListsProps> = ({ context }) => {
  const [lists, setLists] = React.useState<IGraphListInfo[]>([]);
  const [loading, setLoading] = React.useState<boolean>(true);
  const [error, setError] = React.useState<string>('');

  React.useEffect(() => {
    let isMounted: boolean = true;

    const loadLists = async (): Promise<void> => {
      try {
        const client: MSGraphClientV3 = await context.msGraphClientFactory.getClient('3');
        const currentWebUrl = new URL(context.pageContext.web.absoluteUrl);
        const webPath = currentWebUrl.pathname.replace(/\/$/, '');
        const siteResponse = await client
          .api(`/sites/${currentWebUrl.hostname}:${webPath}`)
          .select('id')
          .get();

        const siteId = siteResponse.id as string;
        const listResponse = await client
          .api(`/sites/${siteId}/lists`)
          .select('id,displayName,list')
          .get() as IGraphListResponse;

        if (isMounted) {
          const sortedLists = (listResponse.value || []).slice().sort((a, b) => a.displayName.localeCompare(b.displayName));
          setLists(sortedLists);
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
    return <section className={styles.graphLists}>Loading lists (Graph)...</section>;
  }

  if (error) {
    return <section className={styles.graphLists}>Error: {error}</section>;
  }

  return (
    <section className={styles.graphLists}>
      <h3>Lists via Microsoft Graph</h3>
      <ul>
        {lists.map((list) => (
          <li key={list.id}>
            {list.displayName} (template: {list.list?.template || 'n/a'})
          </li>
        ))}
      </ul>
    </section>
  );
};

export default GraphLists;
