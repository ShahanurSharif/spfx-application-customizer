import { Log } from '@microsoft/sp-core-library';
import * as React from 'react';

import styles from './FcCustomizer.module.scss';

export interface IFcCustomizerProps {
  text: string;
}

const LOG_SOURCE: string = 'FcCustomizer';

export default class FcCustomizer extends React.Component<IFcCustomizerProps> {
  public componentDidMount(): void {
    Log.info(LOG_SOURCE, 'React Element: FcCustomizer mounted');
  }

  public componentWillUnmount(): void {
    Log.info(LOG_SOURCE, 'React Element: FcCustomizer unmounted');
  }

  public render(): React.ReactElement<IFcCustomizerProps> {
    return (
      <div className={styles.fcCustomizer}>
        { this.props.text }
      </div>
    );
  }
}
