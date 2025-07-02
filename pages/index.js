import { useState } from 'react';
import ChatInterface from '../components/ChatInterface';
import Head from 'next/head';

export default function Home() {
  return (
    <>
      <Head>
        <title>צ'אטבוט ל-Google Sheets</title>
        <meta name="description" content="ניהול Google Sheets עם AI" />
        <link rel="icon" href="/favicon.ico" />
        <link href="https://fonts.googleapis.com/css2?family=Assistant:wght@400;700&display=swap" rel="stylesheet" />
        <meta name="viewport" content="width=device-width, initial-scale=1" />
      </Head>
      <ChatInterface />
    </>
  );
} 