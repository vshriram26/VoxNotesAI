// jest-dom adds custom jest matchers for asserting on DOM nodes.
// allows you to do things like:
// expect(element).toHaveTextContent(/react/i)
// learn more: https://github.com/testing-library/jest-dom
import '@testing-library/jest-dom';

Object.defineProperty(HTMLCanvasElement.prototype, 'getContext', {
	value: jest.fn(() => ({
		beginPath: jest.fn(),
		clearRect: jest.fn(),
		closePath: jest.fn(),
		drawImage: jest.fn(),
		lineTo: jest.fn(),
		moveTo: jest.fn(),
		restore: jest.fn(),
		save: jest.fn(),
		setTransform: jest.fn(),
		stroke: jest.fn(),
	})),
});

Object.defineProperty(HTMLCanvasElement.prototype, 'toDataURL', {
	value: jest.fn(() => 'data:image/png;base64,test-ink'),
});
