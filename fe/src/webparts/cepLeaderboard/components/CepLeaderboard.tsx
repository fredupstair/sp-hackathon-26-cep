import * as React from 'react';
import styles from './CepLeaderboard.module.scss';
import type { ICepLeaderboardProps } from './ICepLeaderboardProps';
import * as strings from 'CepLeaderboardWebPartStrings';
import { Spinner, SpinnerSize, Stack, Text } from '@fluentui/react';

export default class CepLeaderboard extends React.Component<ICepLeaderboardProps> {
  public render(): React.ReactElement<ICepLeaderboardProps> {
    const { hasTeamsContext } = this.props;

    return (
      <section className={`${styles.cepLeaderboard} ${hasTeamsContext ? styles.teams : ''}`}>
        {/* TODO: Implement CEP Leaderboard
          - Podium (top 3)
          - Filter tabs: Global / Department / Team
          - Paginated ranked list with search
          - Click on user -> open their dashboard (Other mode)
        */}
        <Stack horizontalAlign="center" tokens={{ padding: 24 }}>
          <Spinner size={SpinnerSize.large} label={strings.Loading} />
          <Text variant="medium">{strings.WebPartTitle}</Text>
        </Stack>
      </section>
    );
  }
}
