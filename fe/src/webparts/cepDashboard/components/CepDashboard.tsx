import * as React from 'react';
import styles from './CepDashboard.module.scss';
import type { ICepDashboardProps } from './ICepDashboardProps';
import * as strings from 'CepDashboardWebPartStrings';
import { Spinner, SpinnerSize, Stack, Text } from '@fluentui/react';

export default class CepDashboard extends React.Component<ICepDashboardProps> {
  public render(): React.ReactElement<ICepDashboardProps> {
    const { hasTeamsContext } = this.props;

    return (
      <section className={`${styles.cepDashboard} ${hasTeamsContext ? styles.teams : ''}`}>
        {/* TODO: Implement CEP Dashboard
          - My mode: points, level, progress bar, usage by app, badges
          - Other mode: aggregated view (points, level, badges only)
        */}
        <Stack horizontalAlign="center" tokens={{ padding: 24 }}>
          <Spinner size={SpinnerSize.large} label={strings.Loading} />
          <Text variant="medium">{strings.WebPartTitle}</Text>
        </Stack>
      </section>
    );
  }
}
