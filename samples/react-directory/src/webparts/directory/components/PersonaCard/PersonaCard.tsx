import * as React from 'react';
import styles from './PersonaCard.module.scss';
import { IPersonaCardProps } from './IPersonaCardProps';
import { IPersonaCardState } from './IPersonaCardState';
import { Log } from '@microsoft/sp-core-library';
import { SPComponentLoader } from '@microsoft/sp-loader';

import {
  Avatar,
  Body1,
  Card,
  Subtitle1,
  Text,
} from '@fluentui/react-components';
import { Call16Filled, Location16Filled } from '@fluentui/react-icons';
import { LIVE_PERSONA_COMPONENT_ID, EXP_SOURCE } from '../../../../constants';

export class PersonaCard extends React.Component<
  IPersonaCardProps,
  IPersonaCardState
> {
  constructor(props: IPersonaCardProps) {
    super(props);
    this.state = { livePersonaCard: undefined, pictureUrl: undefined };
  }

  public async componentDidMount(): Promise<void> {
    const sharedLibrary = await this._loadSPComponentById(LIVE_PERSONA_COMPONENT_ID);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const livePersonaCard: any = sharedLibrary.LivePersonaCard;
    this.setState({ livePersonaCard });
  }

  private _LivePersonaCard(): JSX.Element {
    return React.createElement(
      this.state.livePersonaCard,
      {
        serviceScope: this.props.context.serviceScope,
        legacyUpn: this.props.profileProperties.Email,
        onCardOpen: () => { /* no-op */ },
        onCardClose: () => { /* no-op */ },
      },
      this._PersonaCard()
    );
  }

  private _PersonaCard(): JSX.Element {
    const p = this.props.profileProperties;

    return (
      <Card className={styles.documentCard}>
        <Avatar
          name={p.DisplayName}
          image={{ src: `${p.PictureUrl}` }}
          size={120}
          shape="square"
        />

        <div className={styles.personaDetails}>
          {/* HEADER */}
          <div className={styles.header}>
            <Subtitle1 className={styles.textOverflow}>{p.DisplayName}</Subtitle1>
            <Body1 className={`${styles.others} ${styles.textOverflow}`} style={{ fontWeight: 600 }}>
              {p.Title}
            </Body1>
            <Text className={`${styles.others} ${styles.textOverflow}`}>
              {p.Department}
            </Text>
          </div>

          {/* FOOTER */}
          <div className={styles.footer}>
            {p.WorkPhone ? (
              <div className={styles.others}>
                <Call16Filled style={{ fontSize: '12px' }} />
                <span style={{ marginLeft: 5, fontSize: '12px' }}>{p.WorkPhone}</span>
              </div>
            ) : null}

            {p.Location ? (
              <div className={`${styles.others} ${styles.textOverflow}`}>
                <Location16Filled style={{ fontSize: '12px' }} />
                <span style={{ marginLeft: 5, fontSize: '12px' }}>{p.Location}</span>
              </div>
            ) : null}
          </div>
        </div>
      </Card>
    );
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  private async _loadSPComponentById(componentId: string): Promise<any> {
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const component: any = await SPComponentLoader.loadComponentById(componentId);
      return component;
    } catch (error) {
      Log.error(EXP_SOURCE, error);
      throw new Error(error as any);
    }
  }

  public render(): React.ReactElement<IPersonaCardProps> {
    return (
      <div className={styles.personaContainer}>
        {this.state.livePersonaCard ? this._LivePersonaCard() : this._PersonaCard()}
      </div>
    );
  }
}
