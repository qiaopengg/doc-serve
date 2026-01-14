# doc-serve

流式推送 Word 文档服务

## 功能特性

### 1. 完整 DOCX 块流式推送 ⭐ 推荐

按段落切分文档，每个块都是完整的 DOCX 文件，支持：

- ✅ 无乱码：每个 chunk 都是完整 docx
- ✅ 边收边写：收到一个 chunk 立即写入 WPS
- ✅ 打字机效果：逐段落显示
- ✅ 样式完整：保留所有格式
- ✅ 实时可调：可动态调整速度

**接口**: `GET /api/v1/docs/stream-docx`

### 2. 二进制分片流式推送

**接口**: `GET /api/v1/docs/stream`

### 3. 预览增量帧流 + 最终保真 DOCX

**接口**: 
- `GET /api/v1/docs/frames` - 预览流
- `GET /api/v1/docs/download` - 最终 DOCX

## 快速开始

### 安装依赖

```bash
npm install
```

### 开发模式

```bash
npm run dev --workspace=apps/doc-service
```

服务将在 `http://127.0.0.1:3000` 启动

### 构建

```bash
npm run build --workspace=apps/doc-service
```

### 生产运行

```bash
npm run start --workspace=apps/doc-service
```

### 测试

```bash
npm run test --workspace=apps/doc-service
```

## 项目结构

```
.
├── apps/
│   └── doc-service/          # 文档服务
│       └── src/
│           ├── docStore/     # 文档存储
│           │   ├── docxGenerator.ts  # DOCX 生成器
│           │   ├── fsDocStore.ts
│           │   └── types.ts
│           ├── http/         # HTTP 路由
│           ├── routes/       # 路由处理
│           └── test/         # 测试
└── packages/
    └── doc-core/             # 核心库
        └── src/
            └── stream.ts     # 流式处理
```

## API 文档

详见 [接口文档.md](接口文档.md)

## 环境变量

- `PORT` - 服务端口（默认：3000）
- `HOST` - 监听地址（默认：127.0.0.1）
- `DOCS_DIR` - 文档目录（默认：src）
- `CORS_ORIGIN` - CORS 允许的源（默认：*）

## 技术栈

- Node.js
- TypeScript
- docx - Word 文档生成
- 原生 HTTP 模块（无框架）

## 许可

MIT
