import { act, render, screen } from '@testing-library/react';
import App from './App';

jest.mock('pdfjs-dist/legacy/build/pdf', () => ({
  GlobalWorkerOptions: {},
  getDocument: jest.fn(),
  version: 'test',
}), { virtual: true });

jest.mock('./annotationStorage', () => ({
  getStoredAnnotations: jest.fn().mockResolvedValue(null),
  listStoredAnnotations: jest.fn().mockResolvedValue([]),
  saveStoredAnnotations: jest.fn().mockResolvedValue(null),
}));

test('renders the layered workspace controls', async () => {
  await act(async () => {
    render(<App />);
  });

  expect(screen.getByText(/voxnotes ai/i)).toBeInTheDocument();
  expect(screen.getByRole('heading', { name: /all notes/i })).toBeInTheDocument();
  expect(screen.getByRole('button', { name: /quick note/i })).toBeInTheDocument();
  expect(screen.getByRole('button', { name: /new notebook/i })).toBeInTheDocument();
  expect(screen.getByRole('button', { name: /import file/i })).toBeInTheDocument();
});
