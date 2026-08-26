import * as React from 'react';
import { ICajaConsolidadaKpisProps } from './ICajaConsolidadaKpisProps';
import { IImportesMoneda } from '../models/IImportesMoneda';

const cardStyle: React.CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: 14,
  background: '#FAFAFC',
  border: '1px solid #E3E4E8',
  borderRadius: 12,
  boxShadow: '0 1px 2px rgba(0,0,0,0.04)',
  padding: '16px 18px',
  minHeight: 104
};

const rowStyle: React.CSSProperties = {
  display: 'flex',
  alignItems: 'baseline',
  justifyContent: 'space-between',
  gap: 8
};

const labelStyle: React.CSSProperties = {
  fontSize: 11,
  fontWeight: 600,
  color: '#808493',
  letterSpacing: '0.04em',
  textTransform: 'uppercase'
};

const valueStyle: React.CSSProperties = {
  fontSize: 20,
  fontWeight: 700,
  color: '#232529',
  lineHeight: 1.15
};

const titleStyle: React.CSSProperties = {
  fontSize: 16,
  color: '#666B78',
  lineHeight: 1.25
};

function renderResumenMonedas(totales: IImportesMoneda): React.ReactNode {
  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
      <div style={rowStyle}>
        <span style={labelStyle}>USD</span>
        <span style={valueStyle}>
          {(totales.usd || 0).toLocaleString('en-US', { style: 'currency', currency: 'USD' })}
        </span>
      </div>
      <div style={rowStyle}>
        <span style={labelStyle}>ARS</span>
        <span style={valueStyle}>
          {(totales.ars || 0).toLocaleString('es-MX', { style: 'currency', currency: 'MXN' })}
        </span>
      </div>
    </div>
  );
}

function renderIconCircle(background: string, stroke: string, children: React.ReactNode): React.ReactNode {
  return (
    <div
      style={{
        width: 52,
        height: 52,
        borderRadius: '50%',
        background: background,
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
        flexShrink: 0
      }}
    >
      <svg
        width="26"
        height="26"
        viewBox="0 0 24 24"
        fill="none"
        stroke={stroke}
        strokeWidth="2"
        strokeLinecap="round"
        strokeLinejoin="round"
        aria-hidden="true"
      >
        {children}
      </svg>
    </div>
  );
}

function KpiCard(props: {
  title: string;
  background: string;
  stroke: string;
  icon: React.ReactNode;
  children: React.ReactNode;
}): React.ReactElement {
  return (
    <div style={cardStyle}>
      {renderIconCircle(props.background, props.stroke, props.icon)}
      <div style={{ minWidth: 0, flex: 1, display: 'flex', flexDirection: 'column', gap: 6 }}>
        <div style={titleStyle}>{props.title}</div>
        {props.children}
      </div>
    </div>
  );
}

/**
 * KPIs simplificados de la cabecera de Caja Consolidada.
 */
export default class CajaConsolidadaKpis extends React.Component<ICajaConsolidadaKpisProps> {
  public render(): React.ReactElement<ICajaConsolidadaKpisProps> {
    const { cantidadViajes, totalRecibido, totalEgresos, totalRecupero } = this.props;

    return (
      <div
        style={{
          display: 'grid',
          gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))',
          gap: 16,
          marginBottom: 16
        }}
      >
        <KpiCard
          title="Total recibido"
          background="rgba(16,124,16,0.10)"
          stroke="#107C10"
          icon={
            <>
              <polyline points="22 7 13.5 15.5 8.5 10.5 2 17" />
              <polyline points="16 7 22 7 22 13" />
            </>
          }
        >
          {renderResumenMonedas(totalRecibido)}
        </KpiCard>

        <KpiCard
          title="Egresos"
          background="rgba(209,52,56,0.10)"
          stroke="#D13438"
          icon={
            <>
              <line x1="12" y1="5" x2="12" y2="19" />
              <polyline points="5 12 12 19 19 12" />
            </>
          }
        >
          {renderResumenMonedas(totalEgresos)}
        </KpiCard>

        <KpiCard
          title="Recuperos bancarios"
          background="rgba(136,23,152,0.10)"
          stroke="#881798"
          icon={
            <>
              <rect x="2" y="5" width="20" height="14" rx="2" />
              <line x1="2" y1="10" x2="22" y2="10" />
            </>
          }
        >
          {renderResumenMonedas(totalRecupero)}
        </KpiCard>

        <KpiCard
          title="Cantidad de viajes"
          background="rgba(0,120,212,0.10)"
          stroke="#0078D4"
          icon={
            <>
              <path d="M21 10c0 7-9 13-9 13s-9-6-9-13a9 9 0 0 1 18 0z" />
              <circle cx="12" cy="10" r="3" />
            </>
          }
        >
          <div style={valueStyle}>{(cantidadViajes || 0).toLocaleString('es-AR')}</div>
        </KpiCard>
      </div>
    );
  }
}
