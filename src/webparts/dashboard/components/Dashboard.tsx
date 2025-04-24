import * as React from 'react';
import styles from './Dashboard.module.scss';
import { IDashboardProps } from './IDashboardProps';

const Dashboard: React.FC<IDashboardProps> = (props) => {
  return (
    <div className={styles.container}>
      <aside className={styles.sidebar}>
        <div className={styles.logo}>Brasília Segurança</div>
        <ul className={styles.menu}>
          <li>🏠 Início</li>
          <li>👥 RH</li>
          <li>💼 Comercial</li>
          <li>⚖️ Jurídico</li>
          <li>🛒 Compras</li>
          <li>💻 TI</li>
        </ul>
      </aside>

      <main className={styles.main}>
        <h1 className={styles.welcome}>Bem-vindo, {props.userDisplayName}</h1>
        <div className={styles.grid}>
          <div className={styles.card}>
            <img src="https://source.unsplash.com/featured/?meeting" alt="Evento" />
            <div className={styles.overlay}>Fotos do Último Evento</div>
          </div>
          <div className={styles.card}>
            <img src="https://source.unsplash.com/featured/?security" alt="Segurança" />
            <div className={styles.overlay}>Novos Veículos de Ronda</div>
          </div>
          <div className={styles.card}>
            <img src="https://source.unsplash.com/featured/?teamwork" alt="Colaboradores" />
            <div className={styles.overlay}>Novos Colaboradores</div>
          </div>
          <div className={styles.card}>
            <img src="https://source.unsplash.com/featured/?report" alt="Indicadores" />
            <div className={styles.overlay}>Indicadores 2025</div>
          </div>
        </div>
      </main>
    </div>
  );
};

export default Dashboard;
