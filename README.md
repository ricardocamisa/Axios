# Axios Class for VBA
The Axios class is a VBA project that allows you to make HTTP requests with an easy-to-use interface similar to the Axios library in JavaScript. It provides methods for GET, POST, PUT, DELETE, and PATCH requests and allows you to configure headers, data, and URLs.

## Installation

- Import the `Axios` class module into your VBA project.
- Import the `ConvertJson` module from the [VBA-JSON](https://github.com/VBA-tools/VBA-JSON) project into your VBA project.

## Usage

### Create a request

To create a new `Axios` request, use the `configAxios` method. This method will return a new Axios object with the specified configuration.
