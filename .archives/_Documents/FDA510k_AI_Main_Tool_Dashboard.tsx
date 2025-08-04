import React from 'react';
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, PieChart, Pie, Cell } from 'recharts';
import { Card, CardHeader, CardTitle, CardContent } from '@/components/ui/card';

const FDA510kDashboard = () => {
  // Committee Distribution Data
  const committeeData = [
    { name: 'OR', count: 65, percentage: 41.7 },
    { name: 'NE', count: 23, percentage: 14.7 },
    { name: 'DE', count: 16, percentage: 10.3 },
    { name: 'CV', count: 20, percentage: 12.8 },
    { name: 'GU', count: 12, percentage: 7.7 },
    { name: 'Other', count: 20, percentage: 12.8 }
  ];

  // Geographic Distribution Data
  const geoData = [
    { name: 'California', value: 38, color: '#8884d8' },
    { name: 'Northeast', value: 35, color: '#82ca9d' },
    { name: 'Midwest', value: 28, color: '#ffc658' },
    { name: 'Other US', value: 37, color: '#ff8042' },
    { name: 'International', value: 18, color: '#a4de6c' }
  ];

  // Company Size Data (Verified)
  const companySizeData = [
    // Large Companies ($5B+)
    { 
      category: 'Large',
      companies: [
        { name: 'Medtronic', revenue: 31.2, submissions: 7 },
        { name: 'Stryker', revenue: 19.2, submissions: 12 },
        { name: 'BD', revenue: 19.4, submissions: 2 },
        { name: 'Baxter', revenue: 15.1, submissions: 1 },
        { name: 'Boston Scientific', revenue: 13.1, submissions: 6 },
        { name: 'Smith & Nephew', revenue: 5.2, submissions: 8 }
      ]
    },
    // Mid-Size ($500M-$5B)
    { 
      category: 'Mid-Size',
      companies: [
        { name: 'NuVasive', revenue: 1.2, submissions: 4 },
        { name: 'Globus', revenue: 1.1, submissions: 4 },
        { name: 'Shockwave', revenue: 0.73, submissions: 2 }
      ]
    },
    // Small (<$500M)
    { 
      category: 'Small',
      companies: [
        { name: 'SI-BONE', revenue: 0.128, submissions: 2 },
        { name: 'Treace', revenue: 0.187, submissions: 1 }
      ]
    }
  ];

  return (
    <div className="w-full p-4 space-y-6">
      <Card className="w-full">
        <CardHeader>
          <CardTitle>FDA 510(k) Submission Analysis (N=156)</CardTitle>
        </CardHeader>
        <CardContent className="space-y-8">
          {/* Committee Distribution */}
          <div>
            <h3 className="text-lg font-semibold mb-4">Committee Distribution</h3>
            <div className="h-64">
              <BarChart width={600} height={240} data={committeeData}>
                <CartesianGrid strokeDasharray="3 3" />
                <XAxis dataKey="name" />
                <YAxis />
                <Tooltip 
                  formatter={(value, name) => [`${value} submissions`, name]}
                  labelFormatter={label => `Committee: ${label}`}
                />
                <Bar dataKey="count" fill="#8884d8" />
              </BarChart>
            </div>
          </div>

          {/* Geographic Distribution */}
          <div>
            <h3 className="text-lg font-semibold mb-4">Geographic Distribution</h3>
            <div className="h-64">
              <PieChart width={600} height={240}>
                <Pie
                  data={geoData}
                  dataKey="value"
                  nameKey="name"
                  cx="50%"
                  cy="50%"
                  outerRadius={80}
                  label={({name, value}) => `${name}: ${value}`}
                >
                  {geoData.map((entry, index) => (
                    <Cell key={`cell-${index}`} fill={entry.color} />
                  ))}
                </Pie>
                <Tooltip />
                <Legend />
              </PieChart>
            </div>
          </div>

          {/* Company Size Distribution */}
          <div>
            <h3 className="text-lg font-semibold mb-4">Company Size Distribution (Verified)</h3>
            <div className="h-96">
              {companySizeData.map((category) => (
                <div key={category.category} className="mb-6">
                  <h4 className="font-semibold mb-2">{category.category} Companies</h4>
                  <BarChart width={600} height={category.companies.length * 40} data={category.companies}>
                    <CartesianGrid strokeDasharray="3 3" />
                    <XAxis dataKey="name" />
                    <YAxis />
                    <Tooltip 
                      formatter={(value, name) => [`${value} submissions`, 'Submissions']}
                      labelFormatter={label => `Company: ${label}`}
                    />
                    <Bar dataKey="submissions" fill={
                      category.category === 'Large' ? '#8884d8' :
                      category.category === 'Mid-Size' ? '#82ca9d' :
                      '#ffc658'
                    } />
                  </BarChart>
                </div>
              ))}
            </div>
          </div>
        </CardContent>
      </Card>
    </div>
  );
};

export default FDA510kDashboard;
