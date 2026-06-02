#!/usr/bin/env node
/**
 * Netlify 自动部署脚本
 * 从 GitHub 仓库同步部署到 Netlify
 */

const fs = require('fs');
const path = require('path');
const https = require('https');

// 读取环境变量
function loadEnv() {
    const envPath = path.join(__dirname, '.env');
    const envContent = fs.readFileSync(envPath, 'utf8');
    const env = {};
    
    envContent.split('\n').forEach(line => {
        const match = line.match(/^([^#=]+)=(.*)$/);
        if (match) {
            env[match[1].trim()] = match[2].trim();
        }
    });
    
    return env;
}

// HTTP 请求封装
function request(options, data = null) {
    return new Promise((resolve, reject) => {
        const req = https.request(options, (res) => {
            let body = '';
            res.on('data', chunk => body += chunk);
            res.on('end', () => {
                try {
                    resolve({
                        status: res.statusCode,
                        body: JSON.parse(body)
                    });
                } catch {
                    resolve({ status: res.statusCode, body });
                }
            });
        });
        
        req.on('error', reject);
        
        if (data) {
            req.write(JSON.stringify(data));
        }
        
        req.end();
    });
}

// 获取或创建 Netlify 站点
async function getOrCreateSite(env) {
    console.log('🌐 正在检查 Netlify 站点...');
    
    // 先列出所有站点
    const sites = await request({
        hostname: 'api.netlify.com',
        path: '/api/v1/sites',
        method: 'GET',
        headers: {
            'Authorization': `Bearer ${env.NETLIFY_TOKEN}`,
            'Content-Type': 'application/json'
        }
    });
    
    if (sites.status === 200) {
        // 查找是否已存在同名站点
        const existingSite = sites.body.find(site => 
            site.name === env.REPO_NAME || 
            site.build_settings?.repo_path?.includes(env.REPO_NAME)
        );
        
        if (existingSite) {
            console.log('✅ 找到现有站点:', existingSite.name);
            console.log('🔗 站点地址:', existingSite.ssl_url || existingSite.url);
            return existingSite;
        }
    }
    
    console.log('📦 创建新站点...');
    
    // 创建新站点
    const newSite = await request({
        hostname: 'api.netlify.com',
        path: '/api/v1/sites',
        method: 'POST',
        headers: {
            'Authorization': `Bearer ${env.NETLIFY_TOKEN}`,
            'Content-Type': 'application/json'
        }
    }, {
        name: env.REPO_NAME,
        build_settings: {
            repo_url: `https://github.com/${env.GITHUB_USERNAME}/${env.REPO_NAME}`,
            repo_branch: env.BRANCH_NAME,
            dir: '/',
            cmd: '',
            functions_dir: '',
            env: {}
        }
    });
    
    if (newSite.status === 201) {
        console.log('✅ 站点创建成功:', newSite.body.name);
        console.log('🔗 站点地址:', newSite.body.ssl_url || newSite.body.url);
        return newSite.body;
    } else {
        throw new Error(`创建站点失败: ${JSON.stringify(newSite.body)}`);
    }
}

// 触发部署
async function triggerDeploy(env, siteId) {
    console.log('🚀 正在触发部署...');
    
    const deploy = await request({
        hostname: 'api.netlify.com',
        path: `/api/v1/sites/${siteId}/deploys`,
        method: 'POST',
        headers: {
            'Authorization': `Bearer ${env.NETLIFY_TOKEN}`,
            'Content-Type': 'application/json'
        }
    });
    
    if (deploy.status === 200 || deploy.status === 201) {
        console.log('✅ 部署已触发');
        console.log('📋 部署 ID:', deploy.body.id);
        console.log('🌐 预览地址:', deploy.body.deploy_ssl_url || deploy.body.deploy_url);
        return deploy.body;
    } else {
        throw new Error(`触发部署失败: ${JSON.stringify(deploy.body)}`);
    }
}

// 等待部署完成
async function waitForDeploy(env, siteId, deployId) {
    console.log('⏳ 等待部署完成...');
    
    const maxAttempts = 30;
    let attempts = 0;
    
    while (attempts < maxAttempts) {
        await new Promise(resolve => setTimeout(resolve, 2000));
        
        const deploy = await request({
            hostname: 'api.netlify.com',
            path: `/api/v1/sites/${siteId}/deploys/${deployId}`,
            method: 'GET',
            headers: {
                'Authorization': `Bearer ${env.NETLIFY_TOKEN}`
            }
        });
        
        if (deploy.status === 200) {
            const state = deploy.body.state;
            
            if (state === 'ready') {
                console.log('✅ 部署完成！');
                return deploy.body;
            } else if (state === 'error') {
                throw new Error('部署失败: ' + deploy.body.error_message);
            } else {
                process.stdout.write('.');
            }
        }
        
        attempts++;
    }
    
    throw new Error('部署超时，请手动检查 Netlify 控制台');
}

// 更新 HTML 中的二维码 URL
function updateQRCodeUrl(netlifyUrl) {
    const htmlPath = path.join(__dirname, '客户报价单.html');
    let content = fs.readFileSync(htmlPath, 'utf8');
    
    // 替换 generateQRCode 函数中的 URL 逻辑
    const oldLogic = `// 获取当前页面URL
            const currentUrl = window.location.href;`;
    
    const newLogic = `// 获取当前页面URL
            const currentUrl = '${netlifyUrl}客户报价单.html';`;
    
    if (content.includes(oldLogic)) {
        content = content.replace(oldLogic, newLogic);
        fs.writeFileSync(htmlPath, content);
        console.log('✅ 已更新二维码 URL:', netlifyUrl);
        return true;
    }
    
    return false;
}

// 主函数
async function main() {
    console.log('🚀 开始部署到 Netlify...\n');
    
    try {
        // 1. 加载环境变量
        const env = loadEnv();
        
        if (!env.NETLIFY_TOKEN || env.NETLIFY_TOKEN === 'nfp_xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx') {
            console.error('❌ 请先编辑 .env 文件，填写您的 Netlify Token');
            process.exit(1);
        }
        
        // 2. 获取或创建站点
        const site = await getOrCreateSite(env);
        
        // 3. 触发部署
        const deploy = await triggerDeploy(env, site.id);
        
        // 4. 等待部署完成
        const completedDeploy = await waitForDeploy(env, site.id, deploy.id);
        
        // 5. 更新本地 HTML 的二维码 URL
        const netlifyUrl = site.ssl_url || site.url;
        updateQRCodeUrl(netlifyUrl);
        
        console.log('\n========================================');
        console.log('🎉 Netlify 部署完成！');
        console.log('========================================');
        console.log('🔗 站点地址:', netlifyUrl);
        console.log('🌐 完整链接:', `${netlifyUrl}客户报价单.html`);
        console.log('\n📱 二维码已更新为 Netlify 地址');
        console.log('\n✨ 特点：');
        console.log('  • 全球 CDN 加速，访问更快');
        console.log('  • 自动 HTTPS 证书');
        console.log('  • 每次推送到 GitHub 自动重新部署');
        console.log('========================================\n');
        
    } catch (error) {
        console.error('\n❌ 部署失败:', error.message);
        process.exit(1);
    }
}

main();
